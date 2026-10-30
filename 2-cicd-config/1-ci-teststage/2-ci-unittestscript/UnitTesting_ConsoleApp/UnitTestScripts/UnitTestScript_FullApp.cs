// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     UnitTestScript_FullApp.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2024
// Description:  This script conducts unit testing for a full L5X application file by utilizing the Studio 5000 Logix Designer SDK and Factory Talk Logix Echo SDK.
//               Script outputs: detailed console updates, generated files needed to execute unit testing, & generated excel report detailing test pass/fail info
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using LogixDesigner_ClassLibrary;
using LogixEcho_ClassLibrary;
using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using RockwellAutomation.LogixDesigner.Logging;
using System.Drawing;
using System.Xml.Linq;
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;
using static ConsoleFormatter_ClassLibrary.FileManagement;
using static LogixDesigner_ClassLibrary.LogixDesigner;
using static RockwellAutomation.LogixDesigner.LogixProject;
using static UnitTesting_ConsoleApp.StartUnitTest;

namespace UnitTesting_ConsoleApp.UnitTestScripts
{
    /// <summary>
    /// This class contains the methods and logic to programmatically conduct unit testing for Studio 5000 Logix Designer L5X file applications.
    /// </summary>
    internal class UnitTestScript_FullApp
    {
        // "STATIC VARIABLES" - Use to configure unit test "back-end" setup as desired.
        public static readonly string chassisName = "FullAppUnitTest_Chassis"; /* ---------------- Emulated chassis name. */
        public static readonly bool showFullEventLog = false; /* ----------------------------- Capture and print event logger information to the console.
                                                                                               (Useful during troubleshooting.) */
        public static readonly bool deleteEchoChassis = true; /* ----------------------------- Choose whether to keep or delete emulated chassis (including 
                                                                                               its controllers) at the end of testing. */
        public static readonly bool iExcel_keepACDs = true; /* ------------------------------- Choose whether to keep or delete generated ACD file used during
                                                                                               test execution. */

        /// <summary>
        /// This unit test example has the following steps.<br/>
        /// 1. The "input excel sheet" is parsed. This excel sheet contains the following information:<br/>
        ///    -  The file path of the Studio 5000 Logix Designer L5X application being tested.<br/>
        ///    -  Test cases specifying what inputs to change and what outputs to test (1 test case per excel column).<br/>
        /// 2. Create an emulated controller and chassis using the Echo SDK if one doesn't already exist.<br/>
        /// 3. Go online with the Studio 5000 Logix Designer L5X application and set to Remote Test mode.<br/>
        /// 4. Commence testing. While online with the emulated controller, the LDSDK is used to change the input tags,<br/>
        ///    then verify expected vs. actual output tag results.<br/>
        /// 5. Put unit test results into a worksheet of an excel workbook.<br/>
        ///    If the excel workbook specified in the input command does not yet exist, the workbook is created.<br/>
        ///    If the excel workbook specified in the input command exists, a new worksheet is added to the workbook.<br/>
        ///    (Note for potential future modifications of this unit test script: the output excel sheet containing the results of the<br/>
        ///     unit test was programmatically created and modified at 4 separate locations, specified with the below region name<br/>
        ///     OUTPUT EXCEL REPORT (location #/4 where workbook is updated).
        /// </summary>
        /// <param name="args">
        /// Inputs from UnitTestProgram.
        /// </param>
        /// <returns>True only if at least one test case ran and every test case passed.</returns>
        public static async Task<bool> RunTest(string[] args)
        {
            #region PARSE VARIABLES & INITIALIZE UNIT TEST
            CreateBanner("UNIT TEST INFO");
            // Parse inputs
            string githubPath = args[0];                           // 1st incoming argument = GitHub folder path
            string inputExcelFileName = args[1];                   // 2nd incoming argument = excel file path
            string hash_mostRecentCommit = args[5];                // 6th incoming argument = hash ID from most recent git commit
            string reportAndGeneratedFilesFolderPath = args[8];    // 9th incoming argument = the folder path to the folder storing generated test files
            string inputExcelFilePath = githubPath + @"2-cicd-config\1-ci-teststage\3-ci-inputexcelworkbooks\" + inputExcelFileName;

            // This folder will contain the generated L5X file used during L5X unit testing.
            // (The L5X filetype contains a lot of information used during tag reading/writing.)
            string generatedFilesFolderPath = reportAndGeneratedFilesFolderPath + @"3-generatedfiles\";

            // Variables containing information about the L5X file to test.
            string iExcel_testObjectType;
            string iExcel_testObjectFilePath;

            // Populate the above variables from the input excel file.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputExcelFilePath)))
            {
                ExcelWorksheet inputExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                iExcel_testObjectType = inputExcelWorksheet.Cells[9, 2].Value.ToString()!.Trim()!.ToUpper()!;
                iExcel_testObjectFilePath = githubPath + inputExcelWorksheet.Cells[9, 3].Value.ToString()!.Trim()!;
            }

            // Check if the test object file path starts and ends with quotation marks. If so, remove them.
            if (iExcel_testObjectFilePath.StartsWith('"') && iExcel_testObjectFilePath.EndsWith('"'))
                iExcel_testObjectFilePath = iExcel_testObjectFilePath[1..^1];
            #endregion

            #region STAGING TEST: create emulated controller & chassis -> convert L5X to ACD -> download to ACD -> put controller in test mode
            string l5xFilePath = iExcel_testObjectFilePath;
            string acdFilePath = generatedFilesFolderPath + currentDateTime + "_" + Path.GetFileNameWithoutExtension(l5xFilePath) + ".ACD";

            // Get variables needed to set up unit test. Information retreived from the L5X file specified in the input excel sheet.
            string testObjectName = GetAttributeValue(l5xFilePath, "RSLogix5000Content", "TargetName", false)!;

            // Print message to console about the selected input excel test information.
            string outputTextFilePath = textFileReportPath;
            string outputExcelFilePath = reportAndGeneratedFilesFolderPath + @"2-excelreports\" + currentDateTime + "_" + testObjectName + "_UnitTestReport.xlsx";

            Console.WriteLine("Unit test target object type:".PadRight(40, ' ') + "Full Application (L5X)");
            Console.WriteLine("Input excel workbook file path:".PadRight(40, ' ') + WrapText(inputExcelFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("Output text report path:".PadRight(40, ' ') + WrapText(outputTextFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("Output excel report file path:".PadRight(40, ' ') + WrapText(outputExcelFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("File to be tested:".PadRight(40, ' ') + WrapText(iExcel_testObjectFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("Retain generated ACD file:".PadRight(40, ' ') + iExcel_keepACDs);

            CreateBanner("STAGING UNIT TEST");
            ConsoleMessage("START converting L5X application file to ACD...", "NEWSECTION", false);

            // Open the target object L5X application.
            LogixProject l5xProject = await OpenLogixProjectAsync(l5xFilePath);
            ConsoleMessage($"Opened L5X application file '{l5xFilePath}'.", "STATUS");

            // Convert L5X to ACD.
            await l5xProject.SaveAsAsync(acdFilePath, true);
            ConsoleMessage($"Converted L5X application project at '{l5xFilePath}' to ACD file at '{acdFilePath}'.", "STATUS");

            // Open the generated ACD application to host unit test.
            LogixProject logixProject = await OpenLogixProjectAsync(acdFilePath);
            ConsoleMessage($"Opened ACD application file '{acdFilePath}'.", "STATUS");

            // Capture and print event logger information to the console. (Useful during troubleshooting.)
            if (showFullEventLog)
                logixProject.AddEventHandler(new StdOutEventLogger());

            // Set up emulated controller (based on the specified ACD file path & unit test static variables) if one does not yet exist.
            ConsoleMessage("START setting up Factory Talk Logix Echo emulated controller...", "NEWSECTION");
            string commPath = LogixEchoMethods.CreateChassisFromACD_Async(acdFilePath, chassisName).GetAwaiter().GetResult();
            ConsoleMessage($"Project communication path specified is '{commPath}'.", "STATUS");

            // Change emulated controller mode to program & verify.
            ConsoleMessage("START changing controller to PROGRAM mode...", "NEWSECTION");
            await ChangeControllerMode_Async(commPath, "PROGRAM", logixProject);
            if (ReadControllerMode_Async(commPath, logixProject).GetAwaiter().GetResult() == "PROGRAM")
                ConsoleMessage("SUCCESS changing controller to PROGRAM mode.", "STATUS", false);
            else
                throw new InvalidOperationException("FAILURE changing controller to PROGRAM mode.");

            // Download programmatically generated ACD application to the emulated controller.
            ConsoleMessage("START downloading ACD file...", "NEWSECTION");
            await DownloadProject_Async(commPath, logixProject);
            ConsoleMessage("SUCCESS downloading ACD file.", "STATUS", false);

            // Change emulated controller mode to test & verify.
            ConsoleMessage("START changing controller to TEST mode...", "NEWSECTION");
            await ChangeControllerMode_Async(commPath, "TEST", logixProject);
            if (ReadControllerMode_Async(commPath, logixProject).GetAwaiter().GetResult() == "TEST")
                ConsoleMessage("SUCCESS changing controller to TEST mode.", "STATUS", false);
            else
                throw new InvalidOperationException("FAILURE changing controller to TEST mode.");
            #endregion

            #region COMMENCE TEST: Set & check parameters for each test case from the excel sheet. Results are committed to output excel worksheet.
            ConsoleMessage($"START {testObjectName} unit testing...", "NEWSECTION");
            int testCases = GetRightmostColumnWithData(inputExcelFilePath, 20) - 6; // The number of test cases provided in the input excel workbook.
            int failureCondition = 0; // This variable tracks the number of failed test cases or controller faults.

            // A workbook without test cases verifies nothing, so it cannot pass.
            if (testCases < 1)
            {
                ConsoleMessage($"No test cases found in row 20 of '{inputExcelFileName}'. At least one test case column is required.", "ERROR");
                failureCondition++;
            }

            // Get the Name, DataType, Usage, and XPath components of each tag from the L5X file & put those contents into an array.
            S5kAtomicTag[] testTags = GetTagsForTest(inputExcelFilePath, l5xFilePath)!;

            #region OUTPUT EXCEL REPORT (location 1/4 where workbook is updated): setting up & formatting output excel with banners & row names
            int lowerColumnLimit = 13 + testTags.Length * 3; // Used for excel sheet formatting.

            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputExcelFilePath)))
            {
                // Store all test information in a worksheet with a uniquely time-stamped name.
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"{currentDateTime}_{testObjectName}");

                worksheet.Cells["B2:O6"].Merge = true;
                worksheet.Cells["B2:O6"].Value = $"{testObjectName} Full Application Unit Test Results";
                worksheet.Cells["B2:O6"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["B2:O6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells["B2:O6"].Style.Font.Size = 26;
                worksheet.Cells["B2:O6"].Style.Font.Bold = true;
                worksheet.Cells["B2:O6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells["B2:O6"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                worksheet.Cells["B8"].Value = "Overall Test Result:";
                worksheet.Cells["B8"].Style.Font.Bold = true;
                worksheet.Cells["B8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                worksheet.Cells["C8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                worksheet.Cells["E8:F8"].Merge = true;
                worksheet.Cells["E8:F8"].Value = "Date Test Run:";
                worksheet.Cells["E8:F8"].Style.Font.Bold = true;
                worksheet.Cells["E8:F8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                worksheet.Cells["G8:I8"].Merge = true;
                worksheet.Cells["G8:I8"].Value = testStartTime.ToString();
                worksheet.Cells["G8:I8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                worksheet.Cells["K8:M8"].Merge = true;
                worksheet.Cells["K8:M8"].Value = "Total Test Time (hh:mm:ss):";
                worksheet.Cells["K8:M8"].Style.Font.Bold = true;
                worksheet.Cells["K8:M8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                worksheet.Cells["B10"].Value = "Test Cases:";
                worksheet.Row(10).Style.Font.Bold = true;
                worksheet.Cells["B10"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                for (int i = 0; i < testCases; i++)
                {
                    worksheet.Cells[10, 3 + i].Value = i + 1;
                }
                worksheet.Cells[10, 3, 10, 3 + testCases].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[11, 3, 11, 3 + testCases].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                worksheet.Cells["B11"].Value = "Test Case Result:";
                worksheet.Row(11).Style.Font.Bold = true;
                worksheet.Cells["B11"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                worksheet.Cells["B13"].Value = "Tested Input Parameters:";
                worksheet.Cells["B13"].Style.Font.Bold = true;

                worksheet.Cells[13, 2, lowerColumnLimit, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                int rowNum = 13;
                foreach (var parameter in testTags)
                {
                    if (parameter.Usage != "OUTPUT")
                    {
                        rowNum++;
                        worksheet.Cells[rowNum, 2].Value = parameter.Name;
                    }
                }

                rowNum = rowNum + 2;
                worksheet.Cells[rowNum, 2].Value = "Tested Output Parameters:";
                worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                foreach (var parameter in testTags)
                {
                    if (parameter.Usage != "INPUT")
                    {
                        rowNum++;
                        worksheet.Cells[rowNum, 2].Value = parameter.Name;
                    }
                }

                rowNum = rowNum + 2;
                worksheet.Cells[rowNum, 2].Value = "Actual Output Parameters:";
                worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                foreach (var parameter in testTags)
                {
                    if (parameter.Usage != "INPUT")
                    {
                        rowNum++;
                        worksheet.Cells[rowNum, 2].Value = parameter.Name;
                    }
                }

                worksheet.Column(2).AutoFit();

                package.Save();
            }
            ConsoleMessage($"Done setting up & formatting output excel test report.", "STATUS");
            #endregion

            // Get the values of the "Safe State" test case #0 for each tag.
            Dictionary<string, string> safeStateTestCase = GetExcelTestCaseValues(inputExcelFilePath, 6);

            // Set values to the known safe state, test case #0 of the input excel sheet.
            ConsoleMessage($"\"Safe State\" test case #0 for {testObjectName} tags shown below:", "STATUS");
            for (int i = 0; i < testTags.Length; i++)
                testTags[i].Value = safeStateTestCase[testTags[i].Name!];
            PrintS5kAtomicTags(testTags, false);

            // TEST CASES LOOP: Iterate through each test case provided in the input excel workbook (each column).
            for (int i = 0; i < testCases; i++)
            {
                // Parameters updated/cleared each test case.
                int testNumber = i + 1;                  // The test case currently being tested.
                int inputExcelColumnNum = i + 7;         // The number of the input excel column from which test case values are being obtained.
                int currentColumnNumForOutExcel = i + 3; // Required value for programmatically creating output excel file.
                int testIfFailure = failureCondition;    // testIfFailure used as a comparison value for whether an individual test case failed.

                ConsoleMessage($"START test case {testNumber}/{testCases}...", "NEWSECTION", false);

                // Set values to the known safe state, test case #0 of the input excel sheet.
                ConsoleMessage($"Setting tags to \"Safe State\" test case #0.", "STATUS");
                await SetTagValuesPerTestCase(testTags, safeStateTestCase, logixProject, false);

                // Get the current test case values to be used during testing.
                Dictionary<string, string> currentTestCaseValues = GetExcelTestCaseValues(inputExcelFilePath, inputExcelColumnNum);

                ConsoleMessage($"Setting input tags for test case {testNumber}/{testCases}.", "STATUS");
                // Set input tags based on the current test case values.
                await SetTagValuesPerTestCase(testTags, currentTestCaseValues, logixProject, true);

                // Get the current inputs tag values to verify parameter outputs.
                S5kAtomicTag[]? outputTags = await GetTagValuesPerTestCase(testTags, currentTestCaseValues, logixProject);

                #region OUTPUT EXCEL REPORT (location 2/4 where workbook is updated): test case parameter values from input excel added to output excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputExcelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                    int rowNum = 13;
                    foreach (var tag in testTags)
                    {
                        if (tag.Usage != "OUTPUT")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (tag.Name == kvp.Key)
                                {
                                    rowNum++;
                                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = GetVariableByDataType(kvp.Value, tag.DataType!);
                                    break;
                                }
                            }
                        }
                    }

                    rowNum = rowNum + 2;

                    foreach (var tag in testTags)
                    {
                        if (tag.Usage != "INPUT")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (tag.Name == kvp.Key)
                                {
                                    rowNum++;
                                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = GetVariableByDataType(kvp.Value, tag.DataType!);
                                    break;
                                }
                            }
                        }
                    }

                    rowNum = rowNum + 2;

                    for (int j = 0; j < testTags.Length; j++)
                    {
                        foreach (var tag in outputTags!)
                        {
                            if ((testTags[i].Usage != "INPUT") && (testTags[i].Name == tag.Name))
                            {
                                currentTestCaseValues[testTags[i].Name!] = tag.Value!;
                                break;
                            }
                        }
                    }

                    foreach (var tag in testTags)
                    {
                        if (tag.Usage != "INPUT")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (tag.Name == kvp.Key)
                                {
                                    rowNum++;
                                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = GetVariableByDataType(kvp.Value, tag.DataType!);
                                    break;
                                }
                            }
                        }
                    }

                    worksheet.Column(currentColumnNumForOutExcel).AutoFit();
                    worksheet.Cells[14, currentColumnNumForOutExcel, lowerColumnLimit, currentColumnNumForOutExcel].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    package.Save();
                }
                #endregion

                // OUTPUT PARAMETER LOOP
                for (int j = 0; j < outputTags!.Length; j++)
                {
                    if ((!string.IsNullOrEmpty(currentTestCaseValues[outputTags[j].Name!])) && (outputTags[j].Usage == "OUTPUT"))
                    {
                        failureCondition += TEST_CompareForExpectedValue(outputTags[j].Name!, currentTestCaseValues[outputTags[j].Name!],
                            outputTags[j].Value!, true); // If values not equal, failure condition increased.
                    }
                }

                #region OUTPUT EXCEL REPORT (location 3/4 where workbook is updated): actual output parameter values added to output excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputExcelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                    if (testIfFailure != failureCondition)
                    {
                        worksheet.Cells[11, currentColumnNumForOutExcel].Value = "FAIL";
                        worksheet.Cells[11, currentColumnNumForOutExcel].Style.Font.Color.SetColor(Color.Red);
                    }
                    else
                    {
                        worksheet.Cells[11, currentColumnNumForOutExcel].Value = "PASS";
                        worksheet.Cells[11, currentColumnNumForOutExcel].Style.Font.Color.SetColor(Color.Green);
                    }

                    package.Save();
                }
                ConsoleMessage($"Updated output excel test report with test case {testNumber}/{testCases}.", "STATUS");
                #endregion
            }

            // Based on the unit test result, print a final result message in red or green.
            if (failureCondition > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ConsoleMessage($"{testObjectName} Unit Test Final Result: FAIL | {failureCondition} Issues Encountered", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                ConsoleMessage($"{testObjectName} Unit Test Final Result: PASS", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            #endregion

            #region END TEST: Print final test results & retain/delete generated test components as specified in input excel sheet.
            ConsoleMessage("START retaining or deleting programmatically generated test components...", "NEWSECTION");

            // Based on the static variable, delete or retain the L5X file used during testing.
            if (!iExcel_keepACDs)
            {
                File.Delete(acdFilePath);
                ConsoleMessage($"Deleted '{acdFilePath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained '{acdFilePath}'.", "STATUS");
            }

            // Delete all backup files (ACDs are so frequently generated that there is little value in retaining backups).
            File.Delete(acdFilePath + ".BAK");

            // Delete the generated files folder if it is empty.
            if (!Directory.EnumerateFileSystemEntries(generatedFilesFolderPath).Any())
            {
                Directory.Delete(generatedFilesFolderPath, true);
                ConsoleMessage($"Deleted empty folder '{generatedFilesFolderPath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained folder '{generatedFilesFolderPath}'.", "STATUS");
            }

            await logixProject.GoOfflineAsync(); // Testing is complete. Go offline with the emulated controller.

            // Based on the static variable deleteEchoChassis, keep or delete the Logix Echo chassis (and its controllers) used during testing.
            if (deleteEchoChassis)
            {
                await LogixEchoMethods.DeleteChassis_Async(chassisName);
                ConsoleMessage($"Deleted Logix Echo chassis named '{chassisName}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained Logix Echo chassis named '{chassisName}'.", "STATUS");
            }

            // Compute how long the test took to run. 
            DateTime testEndTime = DateTime.Now;
            TimeSpan testLength = testEndTime.Subtract(testStartTime);
            string formattedTestLength = testLength.ToString(@"hh\:mm\:ss");

            // Update the test-reports-log.xlsx workbook with test report filepaths generated during testing.
            string outExcelTestLogFilePath = reportAndGeneratedFilesFolderPath + "test-reports-log.xlsx";
            string testResult = (failureCondition == 0) ? "PASS" : "FAIL";
            if (File.Exists(outExcelTestLogFilePath))
            {
                AddRowAndPopulateCells(testResult, testStartTime.ToString("MM/dd/yyyy hh:mm tt"), formattedTestLength, hash_mostRecentCommit, outputTextFilePath,
                outputExcelFilePath, outExcelTestLogFilePath);
                ConsoleMessage($"Updated the report log excel file contents at '{outExcelTestLogFilePath}'.", "NEWSECTION");
            }

            // Update the test-reports-log.txt file with test report filepaths generated during testing.
            string outTextTestLogFilePath = reportAndGeneratedFilesFolderPath + "test-reports-log.txt";
            if (!File.Exists(outTextTestLogFilePath))
            {
                string content = "RESULT |  TEST DATE & TIME   | TEST LENGTH |  COMMIT HASH TESTED                     |   TEST REPORT FILE PATH" + Environment.NewLine +
                    "---------------------------------------------------------------------------------------------------------------------------------------------------------";
                File.WriteAllText(outTextTestLogFilePath, content);
            }
            string logLine = testResult.PadRight(7) + "| " + testStartTime.ToString("MM/dd/yyyy hh:mm tt").PadRight(20) + "| " + formattedTestLength.PadRight(12) + "| " +
                hash_mostRecentCommit.PadRight(17) + "| ";
            string newContent = logLine + outputTextFilePath + Environment.NewLine + logLine + outputExcelFilePath;
            AddThirdLineToTextFile(newContent, outTextTestLogFilePath);
            ConsoleMessage($"Updated the report log text file contents in repo at '{outTextTestLogFilePath}'.", "NEWSECTION");

            ConsoleMessage($"Full application unit testing for '{testObjectName}' completed in {formattedTestLength} (HH:mm:ss).", "NEWSECTION");

            #region OUTPUT EXCEL REPORT (location 4/4 where workbook is updated): test length and overall test result added
            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputExcelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                if (failureCondition > 0)
                {
                    worksheet.Cells["C8"].Value = "FAILURE";
                    worksheet.Cells["B2:O6"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                }
                else
                {
                    worksheet.Cells["C8"].Value = "SUCCESS";
                    worksheet.Cells["B2:O6"].Style.Fill.BackgroundColor.SetColor(Color.Green);
                }

                worksheet.Cells["N8"].Value = formattedTestLength;
                worksheet.Cells["N8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                package.Save();
            }
            #endregion

            // Print out final banner..
            CreateBanner("ENDING UNIT TEST");

            // Retain the specified number of generated unit test reports and generated files used to execute unit testing.
            ConsoleMessage("START file management...", "NEWSECTION", false);
            string textReportsFolderPath = reportAndGeneratedFilesFolderPath + @"1-textreports";
            ConsoleMessage($"Set to retain '{numberOfTextReportsToRetain}' text reports at '{textReportsFolderPath}'", "STATUS");
            RetainMostRecentFiles(textReportsFolderPath, numberOfTextReportsToRetain, ".txt");
            string excelReportsFolderPath = reportAndGeneratedFilesFolderPath + @"2-excelreports";
            ConsoleMessage($"Set to retain '{numberOfExcelReportsToRetain}' excel reports at '{excelReportsFolderPath}'", "STATUS");
            RetainMostRecentFiles(excelReportsFolderPath, numberOfExcelReportsToRetain, ".xlsx");
            string temporaryFilesFolderPath = reportAndGeneratedFilesFolderPath + @"3-generatedfiles";
            ConsoleMessage($"Set to retain '{numberOfGeneratedL5XFilesToRetain}' L5X files at '{temporaryFilesFolderPath}'", "STATUS");
            RetainMostRecentFiles(temporaryFilesFolderPath, numberOfGeneratedL5XFilesToRetain, ".L5X");
            ConsoleMessage($"Set to retain '{numberOfGeneratedACDFilesToRetain}' ACD files at '{temporaryFilesFolderPath}'", "STATUS");
            RetainMostRecentFiles(temporaryFilesFolderPath, numberOfGeneratedACDFilesToRetain, ".ACD");
            ConsoleMessage($"Set to retain '{numberOfGeneratedBAKFilesToRetain}' BAK files at '{temporaryFilesFolderPath}'", "STATUS");
            RetainMostRecentFiles(temporaryFilesFolderPath, numberOfGeneratedBAKFilesToRetain, ".BAK");

            // Print out final banner based on test results.
            if (failureCondition > 0)
                CreateBanner("UNIT TEST FINAL RESULT: FAIL");
            else
                CreateBanner("UNIT TEST FINAL RESULT: PASS");

            // Stop logging the console output to the text file.
            StopLogging();

            return failureCondition == 0;
            #endregion
        }

        #region METHODS: L5X Manipulation
        /// <summary>
        /// Programmatically get the XPath of a specified tag from a Studio 5000 Logix Designer L5X file.
        /// </summary>
        /// <param name="l5xPath">The file path to a Studio 5000 Logix Designer L5X file.</param>
        /// <param name="tagName">The target tag to get the XPath of.</param>
        /// <param name="programName">
        /// If the tag is known to be a program scoped, specify the program name here.<br/>
        /// If this input is left blank and the tag name specified exists in multiple programs, the first program listed in the L5X file is used in the XPath. 
        /// </param>
        /// <returns>
        /// The XPath of a specified Studio 5000 Logix Designer tag.<br/>
        /// OR<br/>
        /// A null string if the tag name specified does not exist within the L5X file.
        /// </returns>
        private static string? GetXPathFromL5X(string l5xPath, string tagName, string? programName = null)
        {
            XDocument xDoc = XDocument.Load(l5xPath);

            // PROGRAM SCOPED XPATH SEARCH: Find all Program elements.
            var programElements = xDoc.Descendants("Programs").Elements("Program");

            // Cycle through each Program element.
            foreach (var program in programElements)
            {
                string programNameFromL5X = program.Attribute("Name")!.Value;

                // If default programName input value null, then return the XPath of the first instance of the tag if it exists.
                // OR
                // If programName input value is specified, return the XPath of the tag in that program if it exists.
                if ((programName == null) || (programName == programNameFromL5X))
                {
                    // Find all Tag elements within the current Program.
                    var programTagElements = program.Descendants("Tags").Elements("Tag");

                    // Rotate through the current Program's tag and if the tag exists, return it's XPath.
                    foreach (var tag in programTagElements)
                    {
                        if (tag.Attribute("Name")!.Value == tagName)
                        {
                            return $"Controller/Programs/Program[@Name='{programNameFromL5X}']/Tags/Tag[@Name='{tagName}']";
                        }
                    }
                }
            }

            // CONTROLLER SCOPED XPATH SEARCH: Find all Tag elements.
            var controllerTagElements = xDoc.Descendants("Controller").Elements("Tags").Elements("Tag");

            // Rotate through the controller scoped tags and if the tag exists, return it's XPath.
            foreach (var tag in controllerTagElements)
            {
                string tagNameFromL5X = tag.Attribute("Name")!.Value;

                if (tagName == tagNameFromL5X)
                {
                    return $"Controller/Tags/Tag[@Name='{tagName}']";
                }
            }

            // Print error messages to the console.
            if (programName != null)
                ConsoleMessage($"No tag found named '{tagName}' within the program '{programName}' in the file '{l5xPath}'.", "ERROR");
            else
                ConsoleMessage($"No tag found named '{tagName}' in the file '{l5xPath}'.", "ERROR");

            return null;
        }

        /// <summary>
        /// Programmatically get the Data Type of a specified tag from a Studio 5000 Logix Designer L5X file.
        /// </summary>
        /// <param name="l5xPath">The file path to a Studio 5000 Logix Designer L5X file.</param>
        /// <param name="tagName">The target tag to get the Data Type of.</param>
        /// <param name="programName">
        /// If the tag is known to be a program scoped, specify the program name here.<br/>
        /// If this input is left blank and the tag name specified exists in multiple programs, the first program listed in the L5X file is used to return the Data Type of. 
        /// </param>
        /// <returns>
        /// The Data Type of a specified Studio 5000 Logix Designer tag.<br/>
        /// OR<br/>
        /// A null string if the tag name specified does not exist within the L5X file.
        /// </returns>
        private static string? GetTagDataTypeFromL5X(string l5xPath, string tagName, string? programName = null)
        {
            XDocument xDoc = XDocument.Load(l5xPath);

            // PROGRAM SCOPED TAG SEARCH
            // Find all Program elements.
            var programElements = xDoc.Descendants("Programs").Elements("Program");

            // Cycle through each Program element.
            foreach (var program in programElements)
            {
                string programNameFromL5X = program.Attribute("Name")!.Value;

                // If default programName input value null, then return the Data Type of the first instance of the tag if it exists.
                // OR
                // If programName input value is specified, return the Data Type of the tag in that program if it exists.
                if ((programName == null) || (programName == programNameFromL5X))
                {
                    // Find all Tag elements within the current Program.
                    var programTagElements = program.Descendants("Tags").Elements("Tag");

                    // Rotate through the current Program's tag and if the tag exists, return it's XPath.
                    foreach (var tag in programTagElements)
                    {
                        if (tag.Attribute("Name")!.Value == tagName)
                        {
                            return tag.Attribute("DataType")!.Value;
                        }
                    }
                }
            }

            // CONTROLLER SCOPED TAG SEARCH
            // Find all Tag elements.
            var controllerTagElements = xDoc.Descendants("Controller").Elements("Tags").Elements("Tag");

            // Cycle through each controller scoped tag.
            foreach (var tag in controllerTagElements)
            {
                string tagNameFromL5X = tag.Attribute("Name")!.Value;

                if (tagName == tagNameFromL5X)
                {
                    return tag.Attribute("DataType")!.Value;
                }
            }

            // Print error messages to the console.
            if (programName != null)
                ConsoleMessage($"No tag found named '{tagName}' within the program '{programName}' in the file '{l5xPath}'.", "ERROR");
            else
                ConsoleMessage($"No tag found named '{tagName}' in the file '{l5xPath}'.", "ERROR");

            return null;
        }

        /// <summary>
        /// Using both the input excel sheet and the L5X application file, get the Name, DataType, XPath, and Usage of each tag.
        /// </summary>
        /// <param name="excelFilePath">The excel workbook file path.</param>
        /// <param name="xmlFilePath">The xml file path.</param>
        /// <returns>An array of S5kAtomicTag structure elements containing the information for each tag used during unit testing.</returns>
        private static S5kAtomicTag[]? GetTagsForTest(string excelFilePath, string xmlFilePath)
        {
            S5kAtomicTag[]? returnTags = null;

            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int numberOfParameters = GetPopulatedCellsInColumnCount(excelFilePath, 2) - 6;
                returnTags = new S5kAtomicTag[numberOfParameters];

                for (int i = 0; i < numberOfParameters; i++)
                {
                    int rowNumber = i + 20;
                    var currentTagName = worksheet.Cells[rowNumber, 2].Value?.ToString()!.Trim()!;
                    var currentTagUsage = worksheet.Cells[rowNumber, 3].Value?.ToString()!.Trim()!;

                    if (currentTagUsage.ToUpper() == "I")
                        returnTags[i].Usage = "INPUT";
                    else if (currentTagUsage.ToUpper() == "O")
                        returnTags[i].Usage = "OUTPUT";
                    else
                        ConsoleMessage($"Invalid usage input '{currentTagUsage}' for tag '{currentTagName}'.", "ERROR");

                    returnTags[i].Name = currentTagName;
                    returnTags[i].XPath = GetXPathFromL5X(xmlFilePath, currentTagName);
                    returnTags[i].DataType = GetTagDataTypeFromL5X(xmlFilePath, currentTagName);
                }
            }

            return returnTags;
        }

        /// <summary>
        /// Get the value of an attribute for a specific complex element.
        /// </summary>
        /// <param name="xmlFilePath">The L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element containing the attribute that will have its value returned.</param>
        /// <param name="attributeName">The name of the attribute that will have its value returned.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>The value of an attribute for a specific complex element, or null if the attribute is not found.</returns>
        public static string? GetAttributeValue(string xmlFilePath, string complexElementName, string attributeName, bool printOut)
        {
            XDocument xdoc = XDocument.Load(xmlFilePath);
            XElement? complexElement = xdoc.Descendants(complexElementName).FirstOrDefault();

            if (complexElement != null)
            {
                XAttribute? attribute = complexElement.Attribute(attributeName);
                if (attribute != null)
                {
                    return attribute.Value;
                }
                else if (printOut)
                {
                    ConsoleMessage($"Attribute '{attributeName}' not found in element '{complexElementName}'.", "ERROR");
                }
            }
            else if (printOut)
            {
                ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
            }

            return null;
        }
        #endregion

        #region METHODS: Excel Manipulation
        /// <summary>
        /// Used to get the number of test cases from the input excel sheet.<br/>
        /// Get an integer  
        /// </summary>
        /// <param name="excelFilePath">The file path to the target excel workbook.</param>
        /// <param name="startRow">The row number from which to start checking columns.</param>
        /// <returns>An integer representation of the right-most, populated column in the first worksheet of an excel workbook.</returns>
        public static int GetRightmostColumnWithData(string excelFilePath, int startRow)
        {
            using var package = new ExcelPackage(new FileInfo(excelFilePath));
            var worksheet = package.Workbook.Worksheets[0];
            int rightmostColumn = 0;
            int totalRows = worksheet.Dimension.End.Row;

            for (int row = startRow; row <= totalRows; row++)
            {
                for (int col = worksheet.Dimension.End.Column; col >= 1; col--)
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                    {
                        rightmostColumn = Math.Max(rightmostColumn, col);
                        break;
                    }
                }
            }

            return rightmostColumn;
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
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

        /// <summary>
        /// Collect the values of each tag to be used during a test case from a specifically formatted excel workbook.
        /// </summary>
        /// <param name="excelFilePath">The file path of the excel workbook containing the test case information.</param>
        /// <param name="columnNumber">The column number of a test case in the excel file.</param>
        /// <returns>A dictionary where the Key is a tag name and the Value is a tag value.</returns>
        public static Dictionary<string, string> GetExcelTestCaseValues(string excelFilePath, int columnNumber)
        {
            Dictionary<string, string> returnDictionary = [];

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int numberOfParameters = GetPopulatedCellsInColumnCount(excelFilePath, 2) - 6;
                for (int rowNumber = 20; rowNumber < numberOfParameters + 20; rowNumber++)
                {
                    returnDictionary[worksheet.Cells[rowNumber, 2].Value?.ToString()!.Trim()!] =
                        worksheet.Cells[rowNumber, columnNumber].Value?.ToString()!.Trim()!;
                }
            }

            return returnDictionary;
        }

        /// <summary>
        /// Add a row 7 to a specified excel workbook's first worksheet, then populate it's cells B7:G7 with the specified inputs.
        /// </summary>
        /// <param name="B7">Input a new value for the cell B7.</param>
        /// <param name="C7">Input a new value for the cell C7.</param>
        /// <param name="D7">Input a new value for the cell D7.</param>
        /// <param name="E7">Input a new value for the cell E7.</param>
        /// <param name="F7">Input a new value for the cell F7.</param>
        /// <param name="G7">Input a new value for the cell G7.</param>
        /// <param name="excelFilePath">The file path of the excel workbook containing the worksheet to be modified.</param>
        public static void AddRowAndPopulateCells(string B7, string C7, string D7, string E7, string F7, string G7, string excelFilePath)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                worksheet.InsertRow(7, 1);
                worksheet.Cells["B8:G8"].CopyStyles(worksheet.Cells["B7:G7"]);
                worksheet.Cells["B7"].Value = B7;
                worksheet.Cells["C7"].Value = C7;
                worksheet.Cells["D7"].Value = D7;
                worksheet.Cells["E7"].Value = E7;
                worksheet.Cells["F7"].Value = F7;
                worksheet.Cells["G7"].Value = G7;
                for (int i = 2; i < 8; i++)
                {
                    worksheet.Column(i).AutoFit();
                }
                package.Save();
            }
        }
        #endregion

        #region METHODS: TEST & helper methods
        /// <summary>
        /// A test to compare the expected and actual values of a tag.
        /// </summary>
        /// <param name="tagName">The name of the tag to be tested.</param>
        /// <param name="expectedValue">The expected value of the tag under test.</param>
        /// <param name="actualValue">The actual value of the tag under test.</param>
        /// <returns>Return an integer value 1 for test failure and an integer value 0 for test success.</returns>
        /// <remarks>
        /// The integer output is added to an integer that tracks the total number of failures.<br/>
        /// At the end of all testing, the overall SUCCESS/FAILURE of this CI/CD test stage is determined whether its value is greater than 0.
        /// </remarks>
        private static int TEST_CompareForExpectedValue(string tagName, string expectedValue, string actualValue, bool printOut)
        {
            if (expectedValue != actualValue)
            {
                if (printOut)
                    ConsoleMessage($"{tagName} expected value '{expectedValue}' & actual value '{actualValue}' NOT equal.", "FAIL");

                return 1;
            }
            else
            {
                if (printOut)
                    ConsoleMessage($"{tagName} expected value '{expectedValue}' & actual value '{actualValue}' EQUAL.", "PASS");

                return 0;
            }
        }

        /// <summary>
        /// Rename a specified file.
        /// </summary>
        /// <param name="oldFilePath">A string containing the full file path of the file to be renamed.</param>
        /// <param name="newFileName">A string containing the new name of the file.</param>
        public static void RenameFile(string oldFilePath, string newFileName)
        {
            // Get the directory of the old file
            string directory = Path.GetDirectoryName(oldFilePath)!;

            // Combine the directory, new file name, and extension to form the new file path
            string newFilePath = Path.Combine(directory, newFileName);

            // Rename the file
            File.Move(oldFilePath, newFilePath);
        }
        #endregion
    }
}