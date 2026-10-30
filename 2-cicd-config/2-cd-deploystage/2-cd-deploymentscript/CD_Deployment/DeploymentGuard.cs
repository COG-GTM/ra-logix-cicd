// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     DeploymentGuard.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2026
// Description:  Validates the deployment workbook before any SDK call: target controllers must be on an allowlist kept outside the repository, and project
//               files must resolve inside the checkout's production-files folder.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

namespace CD_Deployment
{
    /// <summary>
    /// Thrown when the deployment must stop before (or instead of) downloading to a controller.
    /// </summary>
    public class DeploymentAbortedException : Exception
    {
        public DeploymentAbortedException(string message) : base(message) { }
    }

    /// <summary>
    /// One validated row of the deployment workbook.
    /// </summary>
    /// <param name="RowNumber">The workbook row the entry was read from.</param>
    /// <param name="ApplicationFilePath">The absolute path to the ACD/L5X file, guaranteed to be inside the deployable folder of the checkout.</param>
    /// <param name="ControllerCommPath">The controller communication path, guaranteed to be on the approved controller allowlist.</param>
    public record DeploymentTarget(int RowNumber, string ApplicationFilePath, string ControllerCommPath);

    /// <summary>
    /// Deployment target validation. Contains no Logix SDK calls so that every check runs before anything touches a controller.
    /// </summary>
    public static class DeploymentGuard
    {
        /// <summary>Environment variable holding the path to the approved controller list (e.g. a Jenkins "Secret file" credential).</summary>
        public const string ApprovedControllersFileVariable = "CD_APPROVED_CONTROLLERS_FILE";

        /// <summary>Folder of the checkout from which application files may be deployed.</summary>
        public const string DeployableFolder = "1-production-files";

        private static readonly string[] AllowedExtensions = { ".ACD", ".L5X" };

        private static readonly StringComparison PathComparison =
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

        /// <summary>
        /// Load the approved controller communication paths from the file named by <see cref="ApprovedControllersFileVariable"/>.<br/>
        /// One communication path per line; blank lines and lines starting with '#' are ignored.
        /// </summary>
        /// <param name="allowlistFilePath">The value of the <see cref="ApprovedControllersFileVariable"/> environment variable.</param>
        /// <param name="checkoutRoot">The repository checkout root. The allowlist must not live inside it.</param>
        /// <returns>The set of approved controller communication paths (compared case-insensitively).</returns>
        /// <exception cref="DeploymentAbortedException">Thrown if the allowlist is missing, inside the checkout, malformed, or empty.</exception>
        public static IReadOnlySet<string> LoadApprovedControllers(string? allowlistFilePath, string checkoutRoot)
        {
            if (string.IsNullOrWhiteSpace(allowlistFilePath))
                throw new DeploymentAbortedException($"Environment variable '{ApprovedControllersFileVariable}' is not set. Configure the approved controller " +
                    "list as a Jenkins credential or node setting outside the repository.");

            string fullAllowlistPath = Path.GetFullPath(allowlistFilePath);
            if (IsInsideFolder(fullAllowlistPath, Path.GetFullPath(checkoutRoot)))
                throw new DeploymentAbortedException($"Approved controller list '{fullAllowlistPath}' is inside the repository checkout. It must be kept outside " +
                    "the repository so that a commit cannot change which controllers the pipeline may deploy to.");

            if (!File.Exists(fullAllowlistPath))
                throw new DeploymentAbortedException($"Approved controller list '{fullAllowlistPath}' does not exist.");

            var approved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string[] lines = File.ReadAllLines(fullAllowlistPath);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0 || line.StartsWith('#'))
                    continue;
                if (line.Any(char.IsControl))
                    throw new DeploymentAbortedException($"Approved controller list line {i + 1} contains control characters.");
                approved.Add(line);
            }

            if (approved.Count == 0)
                throw new DeploymentAbortedException($"Approved controller list '{fullAllowlistPath}' contains no controller communication paths.");

            return approved;
        }

        /// <summary>
        /// Validate one workbook row. Nothing from the row may be used unless this returns.
        /// </summary>
        /// <param name="rowNumber">The workbook row number (used in error messages).</param>
        /// <param name="applicationFileCell">The raw value of the application file path cell, relative to the checkout root.</param>
        /// <param name="controllerCommPathCell">The raw value of the controller communication path cell.</param>
        /// <param name="checkoutRoot">The repository checkout root.</param>
        /// <param name="approvedControllers">The approved controller communication paths.</param>
        /// <returns>The validated deployment target.</returns>
        /// <exception cref="DeploymentAbortedException">Thrown if the controller is not approved or the file is outside the deployable folder.</exception>
        public static DeploymentTarget ValidateRow(int rowNumber, object? applicationFileCell, object? controllerCommPathCell, string checkoutRoot,
            IReadOnlySet<string> approvedControllers)
        {
            string commPath = controllerCommPathCell?.ToString()?.Trim() ?? "";
            if (commPath.Length == 0)
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: controller communication path is empty.");
            if (commPath.Any(char.IsControl) || !approvedControllers.Contains(commPath))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: controller communication path '{Sanitize(commPath)}' is not on the approved " +
                    $"controller list ('{ApprovedControllersFileVariable}').");

            string applicationFilePath = ResolveApplicationFile(rowNumber, applicationFileCell?.ToString()?.Trim() ?? "", checkoutRoot);
            return new DeploymentTarget(rowNumber, applicationFilePath, commPath);
        }

        /// <summary>
        /// Resolve a workbook application file path against the checkout and refuse anything outside <see cref="DeployableFolder"/>.
        /// </summary>
        /// <param name="rowNumber">The workbook row number (used in error messages).</param>
        /// <param name="relativePath">The application file path from the workbook, relative to the checkout root.</param>
        /// <param name="checkoutRoot">The repository checkout root.</param>
        /// <returns>The absolute, normalized application file path.</returns>
        /// <exception cref="DeploymentAbortedException">Thrown if the path is empty, rooted, escapes the deployable folder, is a link, has an unsupported
        /// extension, or does not exist.</exception>
        public static string ResolveApplicationFile(int rowNumber, string relativePath, string checkoutRoot)
        {
            if (relativePath.Length == 0)
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: application file path is empty.");
            if (relativePath.Any(char.IsControl))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: application file path contains control characters.");

            string normalized = relativePath.Replace('\\', Path.DirectorySeparatorChar).Replace('/', Path.DirectorySeparatorChar);
            if (Path.IsPathRooted(normalized) || relativePath.StartsWith('\\') || relativePath.StartsWith('/') || relativePath.Contains(':'))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: application file path '{Sanitize(relativePath)}' must be relative to the " +
                    "repository checkout.");

            string root = Path.GetFullPath(checkoutRoot);
            string deployableRoot = Path.GetFullPath(Path.Combine(root, DeployableFolder));
            string fullPath = Path.GetFullPath(Path.Combine(root, normalized));
            if (!IsInsideFolder(fullPath, deployableRoot))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: application file path '{Sanitize(relativePath)}' resolves outside " +
                    $"'{deployableRoot}'.");

            if (!AllowedExtensions.Any(ext => fullPath.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: application file '{Sanitize(relativePath)}' is not an ACD or L5X file.");

            if (!File.Exists(fullPath))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: application file '{fullPath}' does not exist.");

            // A link could still point outside the deployable folder, so refuse links anywhere from the folder down to the file.
            for (string? current = fullPath; current != null && IsInsideFolder(current, deployableRoot); current = Path.GetDirectoryName(current))
                RejectLink(rowNumber, current);
            RejectLink(rowNumber, deployableRoot);

            return fullPath;
        }

        private static void RejectLink(int rowNumber, string path)
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.ReparsePoint))
                throw new DeploymentAbortedException($"Workbook row {rowNumber}: '{path}' is a link; links are not allowed in deployment paths.");
        }

        private static bool IsInsideFolder(string fullPath, string fullFolderPath)
        {
            string folder = Path.TrimEndingDirectorySeparator(fullFolderPath) + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(folder, PathComparison);
        }

        private static string Sanitize(string value) =>
            new string(value.Select(c => char.IsControl(c) ? '?' : c).ToArray());
    }
}
