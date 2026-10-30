// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     DeploymentProgram.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2024
// Description:  This script programmatically downloads specified ACD applications to specified controllers.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using LogixEcho_ClassLibrary;
using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using RockwellAutomation.LogixDesigner.Logging;
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;
using static LogixDesigner_ClassLibrary.LogixDesigner;
using static RockwellAutomation.LogixDesigner.LogixProject;

namespace CD_Deployment
{
    /// <summary>
    /// This class contains the methods and scripts required to perform scripted PLC downloads.
    /// </summary>
    public class AutomatedDownload
    {
        // "STATIC VARIABLES" - Use to configure downloading setup as desired.
        public static readonly int consoleCharLengthLimit = 110; /* ---------------------------------- The character length limit of each line printed to the console.*/
        public static readonly bool showFullEventLog = false; /* ------------------------------------- Capture and print event logger information to the console.
                                                                                                         (Useful during troubleshooting.) */
        public static readonly bool deleteEchoChassis = true; /* ------------------------------------- Choose whether to keep or delete emulated chassis (including 
                                                                                                         its controllers) at the end of testing. */
        private const string EchoCommPathPrefix = @"EmulateEthernet\"; /* --------------------------- Communication path prefix of Logix Echo emulated controllers. */
        public static readonly DateTime testStartTime = DateTime.Now; /* ----------------------------- The time during which this test was first initiated. 
                                                                                                         (Used at end of test to calculate unit test length.) */
        public static readonly string currentDateTime = testStartTime.ToString("yyyyMMddHHmmss"); /* - Time during which test was first initiated.
                                                                                                         (Used to name generated files & test reports.) */

        /// <summary>
        /// Programmatically download specified ACD file applications to specified controllers.<br/>
        /// The ACD file paths and controller commnication paths are specified in an excel workbook contained within the local GitHub folder (method input arg).
        /// </summary>
        /// <param name="args">
        /// args[0] = The file path to the local GitHub folder (example format: C:\Users\TestUser\Desktop\example-github-repo\).<br/>
        /// args[1] = The folder path to which generated files will be saved during testing (example format: C:\CI-Pipeline-Files\).<br/>
        /// The environment variable CD_APPROVED_CONTROLLERS_FILE must name a file, kept outside the repository, listing the approved controller
        /// communication paths (one per line).
        ///</param>
        /// <returns>A Task that programmatically downloads ACDs to controllers and returns the process exit code (0 = success).</returns>
        static async Task<int> Main(string[] args)
        {
            // Handle incorrect number of arguments console output.
            if (args.Length != 2)
            {
                CreateBanner("INCORRECT NUMBER OF INPUTS");
                Console.WriteLine("Correct Command: ".PadRight(20, ' ') + WrapText(@".\CD_Deployment.exe inputExcelFilePath reportAndGeneratedFilesFolderPath",
                    20, consoleCharLengthLimit));
                Console.WriteLine("Example Format: ".PadRight(20, ' ') + WrapText(@".\CD_Deployment.exe C:\Users\RAUser\Desktop\2-cicd-config\2-cd-deploystage\
                    3-cd-inputexcelworkbooks\DeployToControllersWorkbook.xlsx C:\CI-Pipeline-Files\", 20, consoleCharLengthLimit));
                CreateBanner("END");
                return 2;
            }

            // Parse the input excel sheet needed to determine which modules are to have their firmware verified & flashed if needed.
            string githubPath = args[0];                           // 1st incoming argument = GitHub folder path
            string reportAndGeneratedFilesFolderPath = args[1];    /* 2nd incoming argument = the folder path to the folder storing generated test files
                                                                      Note that this input argument is only included to demo downloading with no hardware present. 
                                                                      This input argument reference can be freely deleted if no emulated controllers are included
                                                                      in the DeployToControllersWorkbook excel workbook list of target download controllers.*/
            string inputExcelFilePath = Path.Combine(githubPath, "2-cicd-config", "2-cd-deploystage", "3-cd-inputexcelworkbooks",
                "DeployToControllersWorkbook.xlsx");

            // Create the local folders that will contain the test reports and generated file contents.
            if (!Directory.Exists(reportAndGeneratedFilesFolderPath))
                Directory.CreateDirectory(reportAndGeneratedFilesFolderPath);

            // Print unit test banner to the console.
            Console.WriteLine("\n  ".PadRight(consoleCharLengthLimit - 2, '='));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            string bannerContents = "CD DEPLOYMENT STAGE | " + DateTime.Now + " " + TimeZoneInfo.Local;
            int padding = (consoleCharLengthLimit - bannerContents.Length) / 2;
            Console.WriteLine(bannerContents.PadLeft(bannerContents.Length + padding).PadRight(consoleCharLengthLimit));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            Console.WriteLine("  ".PadRight(consoleCharLengthLimit - 2, '=') + "\n");

            CreateBanner(".NET INFO");
            Console.WriteLine("LDSDK .NET Core version: ".PadRight(40, ' ') + "8.0");
            Console.WriteLine("EchoSDK .NET Core version: ".PadRight(40, ' ') + "6.0");

            CreateBanner("DEPLOYMENT INFO");
            ConsoleMessage("START downloading specified ACD application files to specified controllers...", "NEWSECTION", false);
            ConsoleMessage($"Input excel workbook used is '{inputExcelFilePath}'.", "STATUS");

            try
            {
                await DeployAsync(githubPath, reportAndGeneratedFilesFolderPath, inputExcelFilePath);
            }
            catch (Exception e) when (e is DeploymentAbortedException || e is InvalidOperationException)
            {
                ConsoleMessage(e.Message, "ERROR");
                CreateBanner("CD DEPLOYMENT ABORTED");
                return 1;
            }

            // Compute how long the test took to run and print final banner.
            DateTime testEndTime = DateTime.Now;
            TimeSpan testLength = testEndTime.Subtract(testStartTime);
            string formattedTestLength = testLength.ToString(@"hh\:mm\:ss");
            CreateBanner($"CD DEPLOYMENT SCRIPT COMPLETED IN {formattedTestLength} (HH:mm:ss)");
            return 0;
        }

        /// <summary>
        /// Validate every row of the deployment workbook, then download each application to its controller.<br/>
        /// No Logix SDK or Echo SDK call is made until all rows have passed validation.
        /// </summary>
        /// <param name="githubPath">The file path to the local GitHub folder (repository checkout root).</param>
        /// <param name="reportAndGeneratedFilesFolderPath">The folder path to which generated files are saved.</param>
        /// <param name="inputExcelFilePath">The deployment workbook file path.</param>
        /// <returns>A Task that programmatically downloads ACDs to controllers.</returns>
        /// <exception cref="DeploymentAbortedException">Thrown if validation fails or a controller is not confirmed in the required mode.</exception>
        private static async Task DeployAsync(string githubPath, string reportAndGeneratedFilesFolderPath, string inputExcelFilePath)
        {
            IReadOnlySet<string> approvedControllers = DeploymentGuard.LoadApprovedControllers(
                Environment.GetEnvironmentVariable(DeploymentGuard.ApprovedControllersFileVariable), githubPath);
            List<DeploymentTarget> targets = ReadDeploymentTargets(inputExcelFilePath, githubPath, approvedControllers);
            ConsoleMessage($"Validated {targets.Count} workbook row(s) against {approvedControllers.Count} approved controller(s).", "STATUS");
            CreateBanner("BEGIN DOWNLOADS");

            // Download ACDs to each controller specified in the input excel workbook.
            for (int i = 0; i < targets.Count; i++)
            {
                string plcCommPath = targets[i].ControllerCommPath;
                string acdFilePath = targets[i].ApplicationFilePath;
                string generatedACD = Path.Combine(reportAndGeneratedFilesFolderPath,
                    currentDateTime + "_" + Path.GetFileNameWithoutExtension(acdFilePath) + ".ACD");

                // If the application file is an L5X file type, convert it to an ACD.
                if ((acdFilePath.EndsWith("L5X", StringComparison.OrdinalIgnoreCase) || acdFilePath.EndsWith("l5x", StringComparison.OrdinalIgnoreCase)) && !File.Exists(generatedACD))
                {
                    if (i == 0)
                    {
                        ConsoleMessage($"Application file path '{acdFilePath}' is an L5X file type. Start converting L5X application file to ACD...", "NEWSECTION", false);
                    }
                    else
                    {
                        ConsoleMessage($"Application file path '{acdFilePath}' is an L5X file type. Start converting L5X application file to ACD...", "NEWSECTION");
                    }

                    // Open the target object L5X application.
                    LogixProject l5xProject = await OpenLogixProjectAsync(acdFilePath);
                    ConsoleMessage($"Opened L5X application file '{acdFilePath}'.", "STATUS");

                    // The file path to the generated ACD file.
                    string l5xFilePath = acdFilePath;
                    acdFilePath = generatedACD;

                    // Convert L5X to ACD.
                    await l5xProject.SaveAsAsync(acdFilePath, true);
                    ConsoleMessage($"Converted L5X application project at '{l5xFilePath}' to ACD file at '{acdFilePath}'.", "STATUS");
                }
                if ((acdFilePath.EndsWith("L5X", StringComparison.OrdinalIgnoreCase) || acdFilePath.EndsWith("l5x", StringComparison.OrdinalIgnoreCase)) && File.Exists(generatedACD))
                    acdFilePath = generatedACD;

                // If one of the input file paths specified are for an emulated controller, create that Echo controller if it did not previously exist.
                bool isEmulated = plcCommPath.StartsWith(EchoCommPathPrefix, StringComparison.OrdinalIgnoreCase);
                bool chassisCreated = false;
                if (isEmulated)
                {

                    if (i == 0)
                        ConsoleMessage($"START creating emulated controller for '{plcCommPath}' in chassis 'GeneratedChassis-{i + 1}'.", "NEWSECTION", false);
                    else
                        ConsoleMessage($"START creating emulated controller for '{plcCommPath}' in chassis 'GeneratedChassis-{i + 1}'.", "NEWSECTION");

                    string justIP = plcCommPath.Substring(EchoCommPathPrefix.Length);
                    chassisCreated = await LogixEchoMethods.CreateChassisFromACD_Async(acdFilePath, justIP, $"GeneratedChassis-{i + 1}");
                    ConsoleMessage($"Done creating emulated controller for '{plcCommPath}' in chassis 'GeneratedChassis-{i + 1}'.", "STATUS");
                    ConsoleMessage($"Note that in a typical CI/CD system, the input excel workbook would not contain any emulated controller communication " +
                        $"paths that need to be emulated. Emulated file paths are included only for demo purposes.", "STATUS");
                }

                try
                {
                    // Open instance of the LogixProject class.
                    if ((i == 0) && !isEmulated)
                        ConsoleMessage($"START opening ACD file '{acdFilePath}' for '{plcCommPath}'.", "NEWSECTION", false);
                    else
                        ConsoleMessage($"START opening ACD file '{acdFilePath}' for '{plcCommPath}'.", "NEWSECTION");

                    LogixProject logixProject = await OpenLogixProjectAsync(acdFilePath);

                    // Capture and print event logger information to the console. (Useful during troubleshooting.)
                    if (showFullEventLog)
                        logixProject.AddEventHandler(new StdOutEventLogger());

                    // Change controller mode to program & verify.
                    ConsoleMessage($"START changing controller '{plcCommPath}' to PROGRAM mode...", "NEWSECTION");
                    await ChangeControllerMode_Async(plcCommPath, "PROGRAM", logixProject);
                    if (await ReadControllerMode_Async(plcCommPath, logixProject) != "PROGRAM")
                        throw new DeploymentAbortedException($"FAILURE changing controller '{plcCommPath}' to PROGRAM mode. Download not attempted.");
                    ConsoleMessage($"SUCCESS changing controller '{plcCommPath}' to PROGRAM mode.", "STATUS", false);

                    // Download the specified ACD application to the controller.
                    ConsoleMessage($"START downloading ACD file to '{plcCommPath}'.", "NEWSECTION");
                    await DownloadProject_Async(plcCommPath, logixProject);
                    ConsoleMessage($"SUCCESS downloading ACD file to '{plcCommPath}'.", "STATUS", false);

                    // Change controller mode to test & verify.
                    ConsoleMessage($"START changing controller '{plcCommPath}' to RUN mode...", "NEWSECTION");
                    await ChangeControllerMode_Async(plcCommPath, "RUN", logixProject);
                    if (await ReadControllerMode_Async(plcCommPath, logixProject) != "RUN")
                        throw new DeploymentAbortedException($"FAILURE changing controller '{plcCommPath}' to RUN mode.");
                    ConsoleMessage($"SUCCESS changing controller '{plcCommPath}' to RUN mode.", "STATUS", false);
                }
                catch when (chassisCreated && deleteEchoChassis)
                {
                    await DeleteChassisAfterFailure_Async($"GeneratedChassis-{i + 1}");
                    throw;
                }

                // Based on the static variable deleteEchoChassis, keep or delete the Logix Echo chassis (and its controllers) used to showcase deployment.
                if (deleteEchoChassis)
                {
                    await LogixEchoMethods.DeleteChassis_Async($"GeneratedChassis-{i + 1}");
                    ConsoleMessage($"Done deleting chassis 'GeneratedChassis-{i + 1}' and its controller '{plcCommPath}'.", "NEWSECTION");
                }
                else
                {
                    ConsoleMessage($"Retaining Logix Echo chassis 'GeneratedChassis-{i + 1}' and its controller '{plcCommPath}'.", "NEWSECTION");
                }
            }
        }

        #region METHODS
        /// <summary>
        /// Best-effort removal of an emulated chassis after a failed deployment, so the original failure is still reported.
        /// </summary>
        /// <param name="chassisName">The name of the Logix Echo chassis to delete.</param>
        private static async Task DeleteChassisAfterFailure_Async(string chassisName)
        {
            try
            {
                await LogixEchoMethods.DeleteChassis_Async(chassisName);
                ConsoleMessage($"Deleted chassis '{chassisName}' after the failed deployment.", "STATUS");
            }
            catch (Exception e)
            {
                ConsoleMessage($"Unable to delete chassis '{chassisName}' after the failed deployment: {e.Message}", "ERROR");
            }
        }

        /// <summary>
        /// Read and validate every controller row of the deployment workbook.
        /// </summary>
        /// <param name="excelFilePath">The deployment workbook file path.</param>
        /// <param name="githubPath">The repository checkout root.</param>
        /// <param name="approvedControllers">The approved controller communication paths.</param>
        /// <returns>The validated deployment targets, in workbook order.</returns>
        /// <exception cref="DeploymentAbortedException">Thrown if the workbook is missing or any row fails validation.</exception>
        private static List<DeploymentTarget> ReadDeploymentTargets(string excelFilePath, string githubPath, IReadOnlySet<string> approvedControllers)
        {
            if (!File.Exists(excelFilePath))
                throw new DeploymentAbortedException($"Input excel workbook '{excelFilePath}' does not exist.");

            int numberOfControllers = GetPopulatedCellsInColumnCount(excelFilePath, 2) - 2;
            var targets = new List<DeploymentTarget>();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                for (int i = 0; i < numberOfControllers; i++)
                {
                    int rowNumber = i + 7;
                    targets.Add(DeploymentGuard.ValidateRow(rowNumber, worksheet.Cells[rowNumber, 2].Value, worksheet.Cells[rowNumber, 3].Value,
                        githubPath, approvedControllers));
                }
            }

            if (targets.Count == 0)
                throw new DeploymentAbortedException($"Input excel workbook '{excelFilePath}' lists no controllers.");
            return targets;
        }

        /// <summary>
        /// In the first worksheet of an Excel workbook, get the number of populated cells in the specified column.
        /// </summary>
        /// <param name="excelFilePath">The excel workbook file path.</param>
        /// <param name="columnNumber">The column in which the populated cell count is derived.</param>
        /// <returns>The number of populated cells in the specified column.</returns>
        private static int GetPopulatedCellsInColumnCount(string excelFilePath, int columnNumber)
        {
            int returnCellCount = 0;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                int maxRowNum = worksheet.Dimension.End.Row;

                for (int row = 1; row <= maxRowNum; row++)
                {
                    var cellValue = worksheet.Cells[row, columnNumber].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        returnCellCount++;
                }
            }
            return returnCellCount;
        }
        #endregion
    }
}