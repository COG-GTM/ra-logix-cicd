// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     DeploymentGuardTests.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2026
// Description:  Tests for the deployment workbook validation in DeploymentGuard.cs.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using CD_Deployment;
using Xunit;

namespace CD_Deployment.Tests
{
    public class DeploymentGuardTests : IDisposable
    {
        private const string ApprovedController = @"RA_EXAMPLE!AB_ETH-1\192.168.1.202\Backplane\6";
        private const string UnlistedController = @"RA_EXAMPLE!AB_ETH-1\10.20.0.5\Backplane\0";

        private readonly string tempRoot;
        private readonly string checkout;
        private readonly string allowlistFile;

        public DeploymentGuardTests()
        {
            tempRoot = Path.Combine(Path.GetTempPath(), "cd-guard-" + Guid.NewGuid().ToString("N"));
            checkout = Path.Combine(tempRoot, "checkout");
            Directory.CreateDirectory(Path.Combine(checkout, "1-production-files", "L5Xs"));
            Directory.CreateDirectory(Path.Combine(checkout, "1-production-files", "ACDs"));
            Directory.CreateDirectory(Path.Combine(checkout, "2-cicd-config"));
            Directory.CreateDirectory(Path.Combine(tempRoot, "old-builds"));
            File.WriteAllText(Path.Combine(checkout, "1-production-files", "L5Xs", "ExampleForCICD_L85E.L5X"), "<l5x/>");
            File.WriteAllText(Path.Combine(checkout, "1-production-files", "ACDs", "notes.txt"), "not a project");
            File.WriteAllText(Path.Combine(checkout, "2-cicd-config", "Elsewhere.L5X"), "<l5x/>");
            File.WriteAllText(Path.Combine(tempRoot, "old-builds", "unreviewed.ACD"), "acd");

            allowlistFile = Path.Combine(tempRoot, "approved-controllers.txt");
            File.WriteAllLines(allowlistFile, new[] { "# Line 2 controllers", "", "  " + ApprovedController + "  ", "EmulateEthernet\\127.0.0.1" });
        }

        public void Dispose() => Directory.Delete(tempRoot, true);

        private IReadOnlySet<string> Approved() => DeploymentGuard.LoadApprovedControllers(allowlistFile, checkout);

        [Fact]
        public void ValidRowResolvesInsideProductionFiles()
        {
            DeploymentTarget target = DeploymentGuard.ValidateRow(7, @"1-production-files\L5Xs\ExampleForCICD_L85E.L5X", ApprovedController, checkout,
                Approved());

            Assert.Equal(Path.Combine(checkout, "1-production-files", "L5Xs", "ExampleForCICD_L85E.L5X"), target.ApplicationFilePath);
            Assert.Equal(ApprovedController, target.ControllerCommPath);
        }

        [Fact]
        public void ControllerMatchIgnoresCaseAndSurroundingWhitespace()
        {
            DeploymentTarget target = DeploymentGuard.ValidateRow(7, "1-production-files/L5Xs/ExampleForCICD_L85E.L5X",
                "  " + ApprovedController.ToLowerInvariant() + " ", checkout, Approved());

            Assert.Equal(ApprovedController.ToLowerInvariant(), target.ControllerCommPath);
        }

        [Fact]
        public void UnlistedControllerIsRejected()
        {
            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7,
                @"1-production-files\L5Xs\ExampleForCICD_L85E.L5X", UnlistedController, checkout, Approved()));
            Assert.Contains("not on the approved controller list", e.Message);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void EmptyControllerIsRejected(string? commPath)
        {
            Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7, @"1-production-files\L5Xs\ExampleForCICD_L85E.L5X", commPath,
                checkout, Approved()));
        }

        [Fact]
        public void ControllerWithControlCharactersIsRejected()
        {
            Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7, @"1-production-files\L5Xs\ExampleForCICD_L85E.L5X",
                ApprovedController + "\n", checkout, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ApprovedController + "\n" }));
        }

        [Theory]
        [InlineData(@"..\old-builds\unreviewed.ACD")]
        [InlineData(@"..\..\old-builds\unreviewed.ACD")]
        [InlineData(@"1-production-files\..\..\old-builds\unreviewed.ACD")]
        [InlineData("1-production-files/../../old-builds/unreviewed.ACD")]
        [InlineData(@"2-cicd-config\Elsewhere.L5X")]
        [InlineData(@"1-production-files-other\Example.L5X")]
        [InlineData(@"1-production-files")]
        public void PathOutsideProductionFilesIsRejected(string applicationFilePath)
        {
            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7, applicationFilePath, ApprovedController, checkout,
                Approved()));
            Assert.Contains("resolves outside", e.Message);
        }

        [Theory]
        [InlineData(@"C:\old-builds\unreviewed.ACD")]
        [InlineData(@"\\server\share\unreviewed.ACD")]
        [InlineData(@"\old-builds\unreviewed.ACD")]
        [InlineData("/tmp/unreviewed.ACD")]
        [InlineData(@"1-production-files\L5Xs\ExampleForCICD_L85E.L5X:stream")]
        public void RootedPathIsRejected(string applicationFilePath)
        {
            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7, applicationFilePath, ApprovedController, checkout,
                Approved()));
            Assert.Contains("must be relative", e.Message);
        }

        [Theory]
        [InlineData(@"1-production-files\ACDs\notes.txt", "not an ACD or L5X")]
        [InlineData(@"1-production-files\ACDs\Missing.ACD", "does not exist")]
        [InlineData("", "is empty")]
        public void InvalidApplicationFileIsRejected(string applicationFilePath, string expectedMessage)
        {
            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7, applicationFilePath, ApprovedController, checkout,
                Approved()));
            Assert.Contains(expectedMessage, e.Message);
        }

        [Fact]
        public void LinkInsideProductionFilesIsRejected()
        {
            string link = Path.Combine(checkout, "1-production-files", "ACDs", "Linked.ACD");
            try
            {
                File.CreateSymbolicLink(link, Path.Combine(tempRoot, "old-builds", "unreviewed.ACD"));
            }
            catch (Exception e) when (e is IOException || e is UnauthorizedAccessException)
            {
                return; // Creating links needs extra privileges on some Windows hosts.
            }

            var ex = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.ValidateRow(7, @"1-production-files\ACDs\Linked.ACD",
                ApprovedController, checkout, Approved()));
            Assert.Contains("is a link", ex.Message);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void MissingAllowlistVariableIsRejected(string? allowlistPath)
        {
            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.LoadApprovedControllers(allowlistPath, checkout));
            Assert.Contains(DeploymentGuard.ApprovedControllersFileVariable, e.Message);
        }

        [Fact]
        public void AllowlistInsideCheckoutIsRejected()
        {
            string inRepo = Path.Combine(checkout, "2-cicd-config", "approved-controllers.txt");
            File.WriteAllText(inRepo, UnlistedController);

            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.LoadApprovedControllers(inRepo, checkout));
            Assert.Contains("inside the repository checkout", e.Message);
        }

        [Fact]
        public void MissingOrEmptyAllowlistIsRejected()
        {
            Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.LoadApprovedControllers(Path.Combine(tempRoot, "missing.txt"), checkout));

            string empty = Path.Combine(tempRoot, "empty.txt");
            File.WriteAllLines(empty, new[] { "# nothing approved", "" });
            var e = Assert.Throws<DeploymentAbortedException>(() => DeploymentGuard.LoadApprovedControllers(empty, checkout));
            Assert.Contains("contains no controller", e.Message);
        }

        [Fact]
        public void AllowlistSkipsCommentsAndBlankLines()
        {
            IReadOnlySet<string> approved = Approved();

            Assert.Equal(2, approved.Count);
            Assert.Contains(ApprovedController, approved);
            Assert.Contains("EmulateEthernet\\127.0.0.1", approved);
        }
    }
}
