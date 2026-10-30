// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     UnitTestProgram.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2024
// Description:  This script conducts unit testing utilizing Studio 5000 Logix Designer SDK and Factory Talk Logix Echo SDK.
//               Valid unit test target objects: Add-On Instructions (AOI) Definition L5X file & full application L5X file
//                     (target object example test scripts coming soon for Program L5X, Routine L5X, and Rung L5X files)
//               Script outputs: detailed console updates, generated files needed to execute unit testing, & generated excel report
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using ConsoleFormatter_ClassLibrary;
using OfficeOpenXml;
using UnitTesting_ConsoleApp.UnitTestScripts;
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;
using static ConsoleFormatter_ClassLibrary.FileManagement;
using static ConsoleFormatter_ClassLibrary.PathSafety;

namespace UnitTesting_ConsoleApp
{
    /// <summary>
    /// This class contains the methods and logic to programmatically conduct unit testing for Studio 5000 Logix Designer Add-On Instructions (AOIs) and full<br/>
    /// application L5X files.
    /// </summary>
    public class StartUnitTest
    {
        // "STATIC VARIABLES" - Use to configure unit test "back-end" setup as desired.
        public static readonly DateTime testStartTime = DateTime.Now;                             /* The time during which this test was first initiated. 
                                                                                                     (Used at end of test to calculate unit test length.) */
        public static readonly string currentDateTime = testStartTime.ToString("yyyyMMddHHmmss"); /* Time during which test was first initiated.
                                                                                                     (Used to name generated files & test reports.) */
        public static readonly int numberOfTextReportsToRetain = 20;          // File management input.
        public static readonly int numberOfExcelReportsToRetain = 20;         // File management input.
        public static readonly int numberOfGeneratedACDFilesToRetain = 20;    // File management input.
        public static readonly int numberOfGeneratedL5XFilesToRetain = 20;    // File management input.
        public static readonly int numberOfGeneratedBAKFilesToRetain = 0;     /* File management input. 
                                                                                 (Recommended to delete all backup files due to GitHub integration). */

        /// <summary>
        /// This unit test example has the following steps.<br/>
        /// 1. The "input excel sheet" is parsed. This excel sheet contains the following information:<br/>
        ///    -  The Studio 5000 component is being unit tested (options: AOI_Definition.L5X, Routine.L5X, Rung.L5X, Program.L5X, Application.L5X).<br/>
        ///    -  The file path of the Studio 5000 component being tested.<br/>
        ///    -  A boolean value whether or not to retain generated ACD files.<br/>
        ///    -  A boolean value whether or not to retain generated L5X files.<br/>
        ///    -  Test cases specifying what inputs to change and what outputs to test (1 test case per excel column).<br/>
        ///    -  The number of controller clock cycles to progress each test case before verifying the outputs.<br/>
        /// 2. Create an emulated controller and chassis using the Echo SDK if one doesn't already exist.<br/>
        /// 3. Stage unit testing environment. Each script in the UnitTestScript has more details about how each target object is set up for testing.
        /// 4. Commence testing. While online with the emulated controller, the LDSDK is used to change the input parameters/tags,<br/>
        ///    then verify expected output vs. actual parameter/tag value results.<br/>
        /// 5. Put unit test results into a text file and a worksheet of an excel workbook.<br/>
        ///    If the excel workbook specified in the input command does not yet exist, the workbook is created.<br/>
        ///    If the excel workbook specified in the input command exists, a new worksheet is added to the workbook.<br/>
        ///    (Note for potential future modifications of this unit test script: the output excel sheet containing the results of the<br/>
        ///     unit test was programmatically created and modified at 4 separate locations of this script.)        
        /// </summary>
        /// <param name="args">
        /// args[0] = The file path to the local GitHub folder (example format: C:\Users\TestUser\Desktop\example-github-repo\).<br/>
        /// args[1] = The name of the excel file that determines what is under development (example format: AOIUnitTestInput_ExampleExcel.xlsx).<br/>
        /// args[2] = The name of the person associated with the most recent git commit (example format: "Allen Bradley").<br/>
        /// args[3] = The email of the person associated with the most recent git commit (example format: exampleemail@rockwellautomation.com).<br/>
        /// args[4] = The message of the most recent git commit (example format: "Added XYZ functionality to #_Valve_Program").<br/>
        /// args[5] = The hash ID of the most recent git commit (example format: 85df4eda88c992a130484515fee4eec63d14913d).<br/>
        /// args[6] = The name of the Jenkins job being run (example format: Jenkins-CICD-Example).<br/>
        /// args[7] = The number of the Jenkins job being run (example format: 218).<br/>
        /// args[8] = The folder path to which generated files will be saved during testing (example format: C:\CI-Pipeline-Files\).
        /// </param>
        /// <returns>A Task that unit tests a specific Studio 5000 Logix Designer component specified within the input excel sheet.</returns>
        static async Task Main(string[] args)
        {
            // Incorrect number of parameters console output.
            if (args.Length != 9)
            {
                CreateBanner("INCORRECT NUMBER OF INPUTS");
                Console.WriteLine("Correct Command: ".PadRight(20, ' ') + WrapText(@".\UnitTesting_ConsoleApp.exe githubPath excelFilename name_mostRecentCommit " +
                                  "email_mostRecentCommit message_mostRecentCommit hash_mostRecentCommit jenkinsJobName jenkinsBuildNumber " +
                                  "reportAndGeneratedFilesFolderPath", 20, consoleCharLengthLimit));
                Console.WriteLine("Example Format: ".PadRight(20, ' ') + WrapText(@".\UnitTesting_ConsoleApp.exe C:\Users\TestUser\Desktop\example-github-repo\ " +
                                  @"excel_filename.xlsx 'Allen Bradley' example@gmail.com 'Most recent commit message insert here' " +
                                  @"287bb2c93a2d1c99143d233fd3ed70cdb997f149 Jenkins-CICD-Example 218 C:\CI-Pipeline-Files\", 20, consoleCharLengthLimit));
                CreateBanner("END");
            }

            // Parse incoming arguments.
            string githubPath = args[0];                           // 1st incoming argument = GitHub folder path
            string inputExcelFileName = args[1];                   // 2nd incoming argument = excel file path
            string name_mostRecentCommit = args[2];                // 3rd incoming argument = name of person assocatied with most recent git commit
            string email_mostRecentCommit = args[3];               // 4th incoming argument = email of person associated with most recent git commit
            string message_mostRecentCommit = args[4];             // 5th incoming argument = message provided in the most recent git commit
            string hash_mostRecentCommit = args[5];                // 6th incoming argument = hash ID from most recent git commit
            string jenkinsJobName = args[6];                       // 7th incoming argument = the Jenkins job name
            string jenkinsBuildNumber = args[7];                   // 8th incoming argument = the Jenkins job build number
            string reportAndGeneratedFilesFolderPath = args[8];    // 9th incoming argument = the folder path to the folder storing generated test files
            string inputExcelFilePath = githubPath + @"2-cicd-config\1-ci-teststage\3-ci-inputexcelworkbooks\" + inputExcelFileName;
            string textFileReportPath = reportAndGeneratedFilesFolderPath + @"1-textreports\" + currentDateTime + "_UnitTestReport.txt";

            // Print out error message if input excel workbook does not exist and end script.
            // A workbook name that is not a plain file name is reported as a failed test further below.
            if (IsPlainFileName(inputExcelFileName) && !File.Exists(inputExcelFilePath))
            {
                ConsoleMessage($"Specified input excel workbook named '{inputExcelFileName}' does not exist at '{inputExcelFilePath}'.", "ERROR");
                return;
            }

            // Create the local folders that will contain the test reports and generated file contents.
            if (!Directory.Exists(reportAndGeneratedFilesFolderPath))
                Directory.CreateDirectory(reportAndGeneratedFilesFolderPath);
            if (!Directory.Exists(reportAndGeneratedFilesFolderPath + @"1-textreports\"))
                Directory.CreateDirectory(reportAndGeneratedFilesFolderPath + @"1-textreports\");
            if (!Directory.Exists(reportAndGeneratedFilesFolderPath + @"2-excelreports\"))
                Directory.CreateDirectory(reportAndGeneratedFilesFolderPath + @"2-excelreports\");
            if (!Directory.Exists(reportAndGeneratedFilesFolderPath + @"3-generatedfiles\"))
                Directory.CreateDirectory(reportAndGeneratedFilesFolderPath + @"3-generatedfiles\");

            // Create new test report file (.txt) using the Console printout.
            StartLogging(textFileReportPath);

            // Fail the test if the input excel workbook name is not a plain file name. A FAIL report is written so that Jenkins does not read an older report.
            try
            {
                ValidateFileName(inputExcelFileName);
            }
            catch (UnsafeInputException ex)
            {
                ConsoleMessage(ex.Message, "ERROR");
                CreateBanner("UNIT TEST FINAL RESULT: FAIL");
                StopLogging();
                return;
            }

            // Print unit test banner to the console.
            Console.WriteLine("\n  ".PadRight(consoleCharLengthLimit - 2, '='));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            string bannerContents = "CI UNIT TESTING STAGE | " + DateTime.Now + " " + TimeZoneInfo.Local;
            int padding = (consoleCharLengthLimit - bannerContents.Length) / 2;
            Console.WriteLine(bannerContents.PadLeft(bannerContents.Length + padding).PadRight(consoleCharLengthLimit));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            Console.WriteLine("  ".PadRight(consoleCharLengthLimit - 2, '=') + "\n\n");

            // Print information to the console.
            CreateBanner("GITHUB & JENKINS INFO");
            Console.WriteLine("Test initiated by: ".PadRight(40, ' ') + name_mostRecentCommit);
            Console.WriteLine("Tester contact information: ".PadRight(40, ' ') + email_mostRecentCommit);
            Console.WriteLine("Git commit hash to be verified: ".PadRight(40, ' ') + hash_mostRecentCommit);
            Console.WriteLine("Git commit message to be verified: ".PadRight(40, ' ') + WrapText(message_mostRecentCommit, 40, 85));
            Console.WriteLine("Jenkins job being executed: ".PadRight(40, ' ') + jenkinsJobName);
            Console.WriteLine("Jenkins job build number: ".PadRight(40, ' ') + jenkinsBuildNumber);
            CreateBanner(".NET INFO");
            Console.WriteLine("LDSDK .NET Core version: ".PadRight(40, ' ') + "8.0");
            Console.WriteLine("EchoSDK .NET Core version: ".PadRight(40, ' ') + "6.0");

            // This variable will be populated from the input excel file and contains the type of target object to be unit tested. 
            string iExcel_testObjectType;

            // Parse excel sheet for test object type.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputExcelFilePath)))
            {
                ExcelWorksheet inputExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                iExcel_testObjectType = inputExcelWorksheet.Cells[9, 2].Value.ToString()!.Trim()!.ToUpper()!;
            }

            // Run the test script for the specific input target object.
            // A workbook or L5X value that points outside the allowed folders fails the test before any file is written.
            try
            {
                if (iExcel_testObjectType == "APPLICATION.L5X")
                    await UnitTestScript_FullApp.RunTest(args);
                else if (iExcel_testObjectType == "AOI_DEFINITION.L5X")
                    await UnitTestScript_AOI.RunTest(args);
                else if (iExcel_testObjectType == "RUNG.L5X")
                    ConsoleMessage("Rung unit test has yet to be developed.", "ERROR");
                else if (iExcel_testObjectType == "ROUTINE.L5X")
                    ConsoleMessage("Routine unit test has yet to be developed.", "ERROR");
                else if (iExcel_testObjectType == "PROGRAM.L5X")
                    ConsoleMessage("Program unit test has yet to be developed.", "ERROR");
                else
                {
                    ConsoleMessage($"Test object type '{iExcel_testObjectType}' not supported. Select either AOI_Definition.L5X, Rung.L5X, Routine.L5X, " +
                        $"Program.L5X, or Application.L5X.", "ERROR");
                }
            }
            catch (UnsafeInputException ex)
            {
                ConsoleMessage(ex.Message, "ERROR");
                CreateBanner("UNIT TEST FINAL RESULT: FAIL");
                StopLogging();
            }
        }
    }
}