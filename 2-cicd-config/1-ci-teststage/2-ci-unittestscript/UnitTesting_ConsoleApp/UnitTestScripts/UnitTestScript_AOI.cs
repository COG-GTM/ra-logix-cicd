// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     UnitTestScript_AOI.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2024
// Description:  This script conducts unit testing for an AOI definition L5X file by utilizing the Studio 5000 Logix Designer SDK and Factory Talk Logix Echo SDK.
//               Script outputs: detailed console updates, generated files needed to execute unit testing, & generated excel report detailing test pass/fail info
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using Google.Protobuf;
using LogixDesigner_ClassLibrary;
using LogixEcho_ClassLibrary;
using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using RockwellAutomation.LogixDesigner.Logging;
using System.Drawing;
using System.Text;
using System.Xml.Linq;
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;
using static ConsoleFormatter_ClassLibrary.FileManagement;
using static LogixDesigner_ClassLibrary.LogixDesigner;
using static RockwellAutomation.LogixDesigner.LogixProject;
using static UnitTesting_ConsoleApp.StartUnitTest;

namespace UnitTesting_ConsoleApp.UnitTestScripts
{
    /// <summary>
    /// This class contains the methods and logic to programmatically conduct unit testing for Studio 5000 Logix Designer Add-On Instructions (AOIs).
    /// </summary>
    internal class UnitTestScript_AOI
    {
        // "STATIC VARIABLES" - Use to configure unit test "back-end" setup as desired.
        public static readonly string chassisName = "AOIUnitTest_Chassis"; // ---------------- Emulated chassis name.
        public static readonly string controllerName = "AOIUnitTest_Controller"; // ---------- Emulated controller name.
        public static readonly string processorType = "1756-L85E"; // ------------------------ The type of emulated controller used to host test.
        public static readonly string taskName_Cont = "T00_StageAOIUnitTesting"; // ---------- Name of the continuous task in the Studio 5000 application.
        public static readonly string taskName_Event = "T01_RunAOIUnitTest"; // -------------- Name of the event task in the Studio 5000 application.
        public static readonly string programName_Cont = "P00_StageAOIUnitTesting"; // ------- Name of the program staging test in the Studio 5000 application.
        public static readonly string programName_Event = "P00_RunAOIUnitTest"; // ----------- Name of the program running test in the Studio 5000 application.
        public static readonly string routineName_Cont = "R00_StageAOIUnitTesting"; // ------- Name of the routine staging test in the Studio 5000 application.
        public static readonly string routineName_Event = "R00_RunAOIUnitTest"; /* ----------- Name of the routine running test with the generated from AOI 
                                                                                               definition and imported to the Studio 5000 application. */
        public static readonly string programName_FaultHandler = "PXX_FaultHandler"; // ------ Name of the fault handling program in the Studio 5k application.
        public static readonly string routineName_FaultHandler = "RXX_FaultHandler"; // ------ Name of the fault handling routine in the Studio 5k application.
        public static readonly bool conversionPrintOut = false; /* --------------------------- Making this true will print to console each step
                                                                                               taken when converting the AOI L5X to a rung L5X. */
        public static readonly bool showFullEventLog = false; /* ----------------------------- Capture and print event logger information to the console.
                                                                                               (Useful during troubleshooting.) */
        public static readonly bool deleteEchoChassis = true; /* ----------------------------- Choose whether to keep or delete emulated chassis (including 
                                                                                               its controllers) at the end of testing. */
        public static readonly bool iExcel_keepL5Xs = false; /* ------------------------------ Choose whether to keep or delete generated L5X files used during
                                                                                               test execution. */
        public static readonly bool iExcel_keepACDs = true; /* ------------------------------- Choose whether to keep or delete generated ACD file used during
                                                                                               test execution. */

        /// <summary>
        /// This unit test example has the following steps.<br/>
        /// 1. The "input excel sheet" is parsed. This excel sheet contains the following information:<br/>
        ///    -  The file path of the Studio 5000 Logix Designer Add-On Instruction definition L5X file path being tested.<br/>
        ///    -  Test cases specifying what inputs to change and what outputs to test (1 test case per excel column).<br/>
        ///    -  The number of controller clock cycles to progress each test case before verifying the outputs.<br/>
        /// 2. Create an emulated controller and chassis using the Echo SDK if one doesn't already exist.<br/>
        /// 3. A Studio 5000 Logix Designer ACD application file is created to host unit testing for L5X test inputs.<br/>
        ///    -  An L5X file containing a fault handler program (contents stored within this c# solution) is converted into an ACD file.<br/>
        ///    -  If testing an AOI definition, the AOI's definition L5X is programmatically converted into a Studio 5000 rung containing a<br/>
        ///    populated instance of the AOI instruction (all required/visible instruction inputs are populated). It is then import to the ACD file.<br/>
        /// 4. Commence testing. While online with the emulated controller, the LDSDK is used to change the input parameters/tags,<br/>
        ///    then verify expected vs. actual output parameter results.<br/>
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
        /// <returns>An asyncronous task that executes unit testing on a Studio 5000 Logix Designer Add-On Instruction.</returns>
        public static async Task RunTest(string[] args)
        {
            #region PARSE VARIABLES & INITIALIZE UNIT TEST
            CreateBanner("UNIT TEST INFO");
            // Parse inputs
            string githubPath = args[0];                           // 1st incoming argument = GitHub folder path
            string inputExcelFileName = args[1];                   // 2nd incoming argument = excel file path
            string hash_mostRecentCommit = args[5];                // 6th incoming argument = hash ID from most recent git commit
            string reportAndGeneratedFilesFolderPath = args[8];    // 9th incoming argument = the folder path to the folder storing generated test files
            string inputExcelFilePath = githubPath + @"2-cicd-config\1-ci-teststage\3-ci-inputexcelworkbooks\" + inputExcelFileName;

            // This folder will contain the generated ACD files and/or L5X files required for AOI unit testing.
            string generatedFilesFolderPath = reportAndGeneratedFilesFolderPath + @"3-generatedfiles\";

            // Variables containing information about the AOI file to test.
            string iExcel_testObjectType;
            string iExcel_testObjectFilePath;

            // Populate the above variables from the input excel file.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputExcelFilePath)))
            {
                ExcelWorksheet inputExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                iExcel_testObjectType = inputExcelWorksheet.Cells[9, 2].Value.ToString()!.Trim()!.ToUpper()!;
                iExcel_testObjectFilePath = inputExcelWorksheet.Cells[9, 3].Value.ToString()!.Trim()!;
            }

            // Check if the test object file path starts and ends with quotation marks. If so, remove them.
            if (iExcel_testObjectFilePath.StartsWith('"') && iExcel_testObjectFilePath.EndsWith('"'))
                iExcel_testObjectFilePath = iExcel_testObjectFilePath[1..^1];
            #endregion

            #region STAGING TEST: create ACD file -> create emulated controller & chassis -> import L5Xs -> download to ACD -> put controller in test 
            string commPath = "";
            string partialL5XprojectFilePath = "";
            string newAOIroutineL5XFilePath = "";
            string acdFilePath = "";
            string testObjectName = "";
            LogixProject logixProject;

            // Get variables needed to set up unit test. Information retreived from the L5X or ACD file specified in the input excel sheet.
            // Keep the test object file inside the GitHub folder and the test object name safe to use in generated file names.
            try
            {
                iExcel_testObjectFilePath = ResolvePathWithinFolder(githubPath, iExcel_testObjectFilePath);
                testObjectName = ValidateLogixComponentName(GetAttributeValue(iExcel_testObjectFilePath, "AddOnInstructionDefinition", "Name", false), "AOI definition Name");
            }
            catch (InvalidDataException ex)
            {
                ConsoleMessage(ex.Message, "ERROR");
                CreateBanner("UNIT TEST FINAL RESULT: FAIL");
                StopLogging();
                return;
            }
            string softwareRevision = GetAttributeValue(iExcel_testObjectFilePath, "RSLogix5000Content", "SoftwareRevision", false)!;

            // Print message to console about the selected input excel test information.
            string outputTextFilePath = reportAndGeneratedFilesFolderPath + @"1-textreports\" + currentDateTime + "_" + testObjectName + "_UnitTestReport.txt";
            string outputExcelFilePath = reportAndGeneratedFilesFolderPath + @"2-excelreports\" + currentDateTime + "_" + testObjectName + "_UnitTestReport.xlsx";

            Console.WriteLine("Unit test target object type:".PadRight(40, ' ') + "AOI Definition (L5X)");
            Console.WriteLine("Input excel workbook file path:".PadRight(40, ' ') + WrapText(inputExcelFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("Output text report path:".PadRight(40, ' ') + WrapText(outputTextFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("Output excel report file path:".PadRight(40, ' ') + WrapText(outputExcelFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("File to be tested:".PadRight(40, ' ') + WrapText(iExcel_testObjectFilePath, 40, consoleCharLengthLimit));
            Console.WriteLine("Retain generated ACD files:".PadRight(40, ' ') + iExcel_keepACDs);
            Console.WriteLine("Retain generated L5X files:".PadRight(40, ' ') + iExcel_keepL5Xs);

            CreateBanner("STAGING UNIT TEST");
            ConsoleMessage("START creating & opening ACD application file to be used during testing...", "NEWSECTION", false);

            /* Create the ACD file to host unit test.
                * (Note that these steps are necessary in order to include a program within 'Controller Fault Handler' in Studio 5k.)
                    Step 1. Get a string containing the xml contents needed to make the basic L5X application file without the test target object.
                    Step 2. Create a newly generated L5X application file.
                    Step 3. Open L5X file using LDSDK (has to be opened for step 4).
                    Step 4. Convert the open L5X file to ACD file using the LDSDK. 
                    Step 5. Open ACD file using LDSDK.*/
            string l5xFileContents = L5XFiles.GetFaultHandlingApplicationL5XContents_ForRoutine(routineName_Cont, routineName_Event, programName_Cont,
                programName_Event, taskName_Cont, taskName_Event, routineName_FaultHandler, programName_FaultHandler,
                controllerName, processorType, softwareRevision);                                              // Step 1: Get L5X contents.
            partialL5XprojectFilePath = generatedFilesFolderPath + currentDateTime + "_" + testObjectName + "_Basic.L5X";
            File.WriteAllText(partialL5XprojectFilePath, l5xFileContents);                                     // Step 2: Generate new L5X file.
            LogixProject logixProjectL5X = await OpenLogixProjectAsync(partialL5XprojectFilePath);             // Step 3: Open L5X file.
            ConsoleMessage($"L5X application file created & opened at '{partialL5XprojectFilePath}'.", "STATUS");
            acdFilePath = generatedFilesFolderPath + currentDateTime + "_" + testObjectName + "_WithAOI.ACD";
            await logixProjectL5X.SaveAsAsync(acdFilePath, true);                                              // Step 4: Convert L5X to ACD.
            logixProject = await OpenLogixProjectAsync(acdFilePath);                                           // Step 5: Open ACD file.
            ConsoleMessage($"ACD application file created & opened at '{acdFilePath}'.", "STATUS");

            // Capture and print event logger information to the console. (Useful during troubleshooting.)
            if (showFullEventLog)
                logixProject.AddEventHandler(new StdOutEventLogger());

            // Set up emulated controller (based on the specified ACD file path & unit test static variables) if one does not yet exist.
            ConsoleMessage("START setting up Factory Talk Logix Echo emulated controller...", "NEWSECTION");
            commPath = LogixEchoMethods.CreateChassisFromACD_Async(acdFilePath, chassisName).GetAwaiter().GetResult();
            ConsoleMessage($"Project communication path specified is '{commPath}'.", "STATUS");

            // Finish setting up ACD application for unit testing by importing the AOI & AOI rung L5X files.
            ConsoleMessage("START preparing ACD application environment for unit test...", "NEWSECTION");
            string xPath_aoiDef = @"Controller/AddOnInstructionDefinitions";
            await logixProject.PartialImportFromXmlFileAsync(xPath_aoiDef, iExcel_testObjectFilePath,          // Import the AOI.L5X being tested
                ImportCollisionOptions.OverwriteOnColl);                                                       // to the open ACD application.
            await logixProject.SaveAsync();
            ConsoleMessage($"Imported '{iExcel_testObjectFilePath}' to '{acdFilePath}'.", "STATUS");

            // Convert a copy of the AOIDefinition.L5X into routine.L5X format, then import into the ACD application. 
            // The ladder logic rung contains an instance of the AOI instruction populated with any visible and/or required tags.
            ConsoleMessage($"Print STATUS messages for AOI.L5X to rung.L5X conversion? Currently set to '{conversionPrintOut}'. To change, change the " +
                $"value of the 'conversionPrintOut' static variable in the 'UnitTestScript_AOI.cs' file.", "STATUS");
            newAOIroutineL5XFilePath = CopyFile(iExcel_testObjectFilePath, generatedFilesFolderPath);
            ConvertAOIL5X_DEFINITIONtoROUTINE(newAOIroutineL5XFilePath, routineName_Event, programName_Event, controllerName, conversionPrintOut);

            string xPath_convertedRungFromAOI = $"Controller/Programs/Program[@Name='{programName_Event}']";
            await logixProject.PartialImportFromXmlFileAsync(xPath_convertedRungFromAOI,                       // Import the programmatically created
                newAOIroutineL5XFilePath, ImportCollisionOptions.OverwriteOnColl);                             // rung to the open ACD application.                                                        
            await logixProject.SaveAsync();
            ConsoleMessage($"Imported '{newAOIroutineL5XFilePath}' to '{acdFilePath}'.", "STATUS");

            // Change emulated controller mode to program & verify.
            ConsoleMessage("START changing controller to PROGRAM mode...", "NEWSECTION");
            await ChangeControllerMode_Async(commPath, "PROGRAM", logixProject);
            if (ReadControllerMode_Async(commPath, logixProject).GetAwaiter().GetResult() == "PROGRAM")
                ConsoleMessage("SUCCESS changing controller to PROGRAM mode.", "STATUS", false);
            else
                ConsoleMessage("FAILURE changing controller to PROGRAM mode.", "ERROR", false);

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
                ConsoleMessage("FAILURE changing controller to TEST mode.", "ERROR", false);

            string fullL5XprojectFilePath = generatedFilesFolderPath + currentDateTime + "_" + testObjectName + "_FullProj.L5X";
            await logixProject.SaveAsAsync(fullL5XprojectFilePath, true);
            #endregion

            #region COMMENCE TEST: Set & check parameters for each test case from the excel sheet. Results are committed to output excel worksheet.
            CreateBanner($"START {testObjectName} UNIT TEST");

            // Get the Name, DataType, Usage, Required, and Visible components of each parameter from the AOI definition XML file
            // & put those contents into an array.
            AOIParameter[] AOIParameters = GetParametersFromL5X(iExcel_testObjectFilePath)!;

            // Store the XPath of the AOI Studio 5000 Logix Designer tag that was programmatically created and used during testing.
            string aoiTagXPath = $"Controller/Tags/Tag[@Name='AOI_{testObjectName}']";

            // Unit test variables
            S5kAtomicTag AT_FaultType;         // AT_FaultType tag storing the controller fault type information.
            S5kAtomicTag AT_FaultCode;         // AT_FaultCode tag storing the controller fault code information.
            bool faultedState = false;         // An indicator of whether the controller is faulted or not.
            bool breakOutputParameterLoop;     // Used to break the "OUTPUT PARAMETER LOOP" if controller faulted when setting inputs. 
            int testCases = GetPopulatedCellsInRowCount(inputExcelFilePath, 18) - 2; // The number of test cases provided in the excel input workbook.
            int failureCondition = 0;          // This variable tracks the number of failed test cases or controller faults.
            string previousTestEnableIn = "0"; /* Track the previous value of the EnableIn parameter.
                                                  Used in logic determining whether or not the tag AT_EnableIn needs to be updated. */

            #region OUTPUT EXCEL REPORT (location 1/4 where workbook is updated): setting up & formatting output excel with banners & row names
            int lowerColumnLimit = 13 + AOIParameters.Length * 3; // Used for excel sheet formatting.

            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputExcelFilePath)))
            {
                // Store all test information in a worksheet with a uniquely time-stamped name.
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"{currentDateTime}_{testObjectName}");

                worksheet.Cells["B2:O6"].Merge = true;
                worksheet.Cells["B2:O6"].Value = $"{testObjectName} AOI Unit Test Results";
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
                foreach (var parameter in AOIParameters)
                {
                    if (parameter.Usage != "Output")
                    {
                        rowNum++;
                        worksheet.Cells[rowNum, 2].Value = parameter.Name;
                    }
                }

                rowNum = rowNum + 2;
                worksheet.Cells[rowNum, 2].Value = "Tested Output Parameters:";
                worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                foreach (var parameter in AOIParameters)
                {
                    if (parameter.Usage != "Input")
                    {
                        rowNum++;
                        worksheet.Cells[rowNum, 2].Value = parameter.Name;
                    }
                }

                rowNum = rowNum + 2;
                worksheet.Cells[rowNum, 2].Value = "Actual Output Parameters:";
                worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                foreach (var parameter in AOIParameters)
                {
                    if (parameter.Usage != "Input")
                    {
                        rowNum++;
                        worksheet.Cells[rowNum, 2].Value = parameter.Name;
                    }
                }

                rowNum = rowNum + 2;
                worksheet.Cells[rowNum, 2].Value = "Controller Fault Info:";
                worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                rowNum++;
                worksheet.Cells[rowNum, 2].Value = "Type";
                rowNum++;
                worksheet.Cells[rowNum, 2].Value = "Code";

                worksheet.Column(2).AutoFit();

                package.Save();
            }
            ConsoleMessage($"Done setting up & formatting output excel test report.", "STATUS");
            #endregion

            // Get the values of the "Safe State" test case #0 for each AOI parameter.
            Dictionary<string, string> safeStateTestCase = GetExcelTestCaseValues(inputExcelFilePath, 6);

            // Set values to the known safe state, test case #0 of the input excel sheet.
            await SetMultipleAOIParamVals_Async(aoiTagXPath, safeStateTestCase, AOIParameters, OperationMode.Online, logixProject);
            await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);
            ConsoleMessage($"{testObjectName} parameter values set to \"Safe State\" test case #0, shown below:", "STATUS");
            PrintAOIParameters(AOIParameters, false);

            // TEST CASES LOOP: Iterate through each test case provided in the input excel workbook (each column).
            for (int i = 0; i < testCases; i++)
            {
                // Parameters updated/cleared each test case.
                int testNumber = i + 1;                  // The test case currently being tested.
                int inputExcelColumnNum = i + 7;        // The number of the input excel column from which test case values are being obtained.
                int faultType = 0;                       // Integer variable storing the controller fault type number (used in output excel).
                int faultCode = 0;                       // Integer variable storing the controller fault code number (used in output excel).
                breakOutputParameterLoop = false;        // Used to break the "OUTPUT PARAMETER LOOP" if controller faulted when setting inputs.
                int currentColumnNumForOutExcel = i + 3; // Required value for programmatically creating output excel file.
                int testIfFailure = failureCondition;    // testIfFailure used as a comparison value for whether an individual test case failed.
                int numberOfClockCycles = 0;             // The number of times the controller's clock cycle will be iterated for the current test case.

                ConsoleMessage($"START test case {testNumber}/{testCases}...", "NEWSECTION", false);

                // Set values to the known safe state, test case #0 of the input excel sheet.
                await SetMultipleAOIParamVals_Async(aoiTagXPath, safeStateTestCase, AOIParameters, OperationMode.Online, logixProject);

                // The EnableIn parameter is modified by changing the boolean tag AT_EnableIn (within an XIC instruction before the AOI instruction).
                if (previousTestEnableIn == "0" && safeStateTestCase["EnableIn"] == "1")
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "1", OperationMode.Online, DataType.BOOL, logixProject);
                else if (previousTestEnableIn == "1" && safeStateTestCase["EnableIn"] == "0")
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "0", OperationMode.Online, DataType.BOOL, logixProject);

                // Forward the test by 1 controller clock cycle.
                // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                // The JSR runs the programmatically generated routine containing the AOI instruction.)
                await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);
                ConsoleMessage($"Parameters set to \"Safe State\" test case #0.", "STATUS");

                // Get the current test case values to be used during testing.
                Dictionary<string, string> currentTestCaseValues = GetExcelTestCaseValues(inputExcelFilePath, inputExcelColumnNum);

                // The EnableIn parameter is modified by changing the boolean tag AT_EnableIn (within an XIC instruction before the AOI instruction).
                if (safeStateTestCase["EnableIn"] == "0" && currentTestCaseValues["EnableIn"] == "1")
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "1", OperationMode.Online, DataType.BOOL, logixProject);
                else if (safeStateTestCase["EnableIn"] == "1" && currentTestCaseValues["EnableIn"] == "0")
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "0", OperationMode.Online, DataType.BOOL, logixProject);

                #region UNIT TEST: change all AOI input parameters for the given test case
                ConsoleMessage($"Setting input values for test case {testNumber}/{testCases}.", "STATUS");

                // Set input parameters based on the current test case values.
                await SetMultipleAOIParamVals_Async(aoiTagXPath, currentTestCaseValues, AOIParameters, OperationMode.Online, logixProject, true);

                // Get the number of clock cycles to test for this suite of test values.
                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputExcelFilePath)))
                {
                    ExcelWorksheet inputExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                    numberOfClockCycles = Convert.ToInt32(inputExcelWorksheet.Cells[18, inputExcelColumnNum].Value.ToString()!.Trim()!.ToUpper()!);
                }

                // ITERATE THROUGH CLOCK CYCLES LOOP
                for (int j = 1; j < numberOfClockCycles + 1; j++)
                {
                    // Forward the test by the number of specified controller clock cycle.
                    // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                    // The JSR runs the programmatically generated routine containing the AOI instruction.)
                    ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject).GetAwaiter().GetResult();
                    ConsoleMessage($"Progressed controller by 1 clock cycle ({j}/{numberOfClockCycles}).", "STATUS");

                    // Check if changing the input parameters for this test case caused a controller fault.
                    AT_FaultType = GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultType']", DataType.DINT, logixProject).GetAwaiter().GetResult();
                    AT_FaultCode = GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultCode']", DataType.DINT, logixProject).GetAwaiter().GetResult();
                    faultType = int.Parse(AT_FaultType.Value!); // Used in output excel sheet report.
                    faultCode = int.Parse(AT_FaultCode.Value!); // Used in output excel sheet report.
                    faultedState = AT_FaultType.Value! != "0" || AT_FaultCode.Value! != "0";

                    if (faultedState)
                    {
                        ConsoleMessage($"Controller faulted with type #{AT_FaultType.Value} & code #{AT_FaultCode.Value}.", "ERROR");
                        break; // Stop the ITERATE THROUGH CLOCK CYCLES LOOP if controller faulted.
                    }
                }

                // If controller faulted, attempt to clear it.
                if (faultedState)
                {
                    failureCondition++;

                    ConsoleMessage($"Attempting to clear fault. Setting all input parameter values to the \"Safe State\" test case #0 & verifying if " +
                        $"controller no longer faulted.", "STATUS");

                    // Clear the fault (toggle XIC to clear the two tags AT_FaultType & AT_FaultCode in the Studio 5000 Logix Designer Project).
                    await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ClearFault']", true, OperationMode.Online, logixProject);

                    // Set values to "Safe State" test case #0.
                    await SetMultipleAOIParamVals_Async(aoiTagXPath, safeStateTestCase, AOIParameters, OperationMode.Online, logixProject);

                    // Forward the test by 1 controller clock cycle.
                    // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                    // The JSR runs the programmatically generated routine containing the AOI instruction.)
                    ConsoleMessage($"Progressing controller by 1 clock cycle.", "STATUS");
                    await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);

                    // Check if controller still faulted after setting to "Safe State" test case #0.
                    AT_FaultType = await GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultType']", DataType.DINT, logixProject);
                    AT_FaultCode = await GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultCode']", DataType.DINT, logixProject);
                    faultedState = AT_FaultType.Value! != "0" || AT_FaultCode.Value! != "0";

                    if (faultedState)
                    {
                        ConsoleMessage("Controller still faulted. Ending Test.", "ERROR");
                        break; // Break the "TEST CASES LOOP" 
                    }
                    else if (testNumber < testCases) // Controller is not faulted and there are more tests remaining.
                    {
                        ConsoleMessage($"Fault cleared. Moving to next test case...", "SUCCESS");
                        breakOutputParameterLoop = true; // Break the "OUTPUT PARAMETER LOOP"
                    }
                }
                #endregion

                // Get the current parameter values of the Studio 5000 AOI tag with which to verify parameter outputs.
                ByteString aoiByteString = await logixProject.GetTagValueAsync(aoiTagXPath, OperationMode.Online, DataType.BYTE_ARRAY);
                AOIParameter[] AOIParameters_WithOutputs = GetAOIParameterValues(AOIParameters, aoiByteString);

                #region OUTPUT EXCEL REPORT (location 2/4 where workbook is updated): test case parameter values from input excel added to output excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputExcelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                    int rowNum = 13;
                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Output")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (parameter.Name == kvp.Key)
                                {
                                    rowNum++;
                                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = GetVariableByDataType(kvp.Value, parameter.DataType!);
                                    break;
                                }
                            }
                        }
                    }

                    rowNum = rowNum + 2;

                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Input")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (parameter.Name == kvp.Key)
                                {
                                    rowNum++;
                                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = GetVariableByDataType(kvp.Value, parameter.DataType!);
                                    break;
                                }
                            }
                        }
                    }

                    rowNum = rowNum + 2;

                    foreach (var parameter in AOIParameters_WithOutputs)
                    {
                        if (parameter.Usage != "Input")
                        {
                            rowNum++;
                            worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = GetVariableByDataType(parameter.Value!, parameter.DataType!);
                        }
                    }

                    rowNum = rowNum + 3;
                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = faultType;
                    rowNum++;
                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = faultCode;

                    worksheet.Column(currentColumnNumForOutExcel).AutoFit();
                    worksheet.Cells[14, currentColumnNumForOutExcel, lowerColumnLimit, currentColumnNumForOutExcel].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    package.Save();
                }
                #endregion

                // OUTPUT PARAMETER LOOP
                foreach (var kvp in currentTestCaseValues)
                {
                    if (breakOutputParameterLoop)
                        break;

                    if (GetAOIParameterComponentValue(kvp.Key, "Usage", AOIParameters) != "Input")
                    {
                        string outputValue = GetAOIParameterComponentValue(kvp.Key, "Value", AOIParameters_WithOutputs);
                        failureCondition += TEST_CompareForExpectedValue(kvp.Key, kvp.Value, outputValue, true); // If values not equal, failure condition increased.
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

                // Used in logic determining whether or not the tag AT_EnableIn needs to be updated.
                previousTestEnableIn = currentTestCaseValues["EnableIn"];
            }

            // Based on the AOI unit test result, print a final result message in red or green.
            if (failureCondition == 1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ConsoleMessage($"{testObjectName} Unit Test Final Result: FAIL | {failureCondition} Issue Encountered", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else if (failureCondition > 1)
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

            // Based on the AOI Excel Worksheet for this AOI, keep or delete generated L5X files.
            if (!iExcel_keepL5Xs)
            {
                File.Delete(partialL5XprojectFilePath);
                File.Delete(newAOIroutineL5XFilePath);
                File.Delete(fullL5XprojectFilePath);
                ConsoleMessage($"Deleted '{partialL5XprojectFilePath}'.", "STATUS");
                ConsoleMessage($"Deleted '{newAOIroutineL5XFilePath}'.", "STATUS");
                ConsoleMessage($"Deleted '{fullL5XprojectFilePath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained '{partialL5XprojectFilePath}'.", "STATUS");
                ConsoleMessage($"Retained '{newAOIroutineL5XFilePath}'.", "STATUS");
                ConsoleMessage($"Retained '{fullL5XprojectFilePath}'.", "STATUS");
            }

            // Based on the AOI Excel Worksheet for this AOI, keep or delete the generated ACD file.
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
                ConsoleMessage($"Deleted Logix Echo chassis named '{chassisName}' & controller named '{controllerName}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained Logix Echo chassis named '{chassisName}' & controller named '{controllerName}'.", "STATUS");
            }

            // Compute how long the test took to run. 
            DateTime testEndTime = DateTime.Now;
            TimeSpan testLength = testEndTime.Subtract(testStartTime);
            string formattedTestLength = testLength.ToString(@"hh\:mm\:ss");

            // Update the test-reports-log.xlsx workbook with test report filepaths generated during testing.
            string outExcelTestLogFilePath = reportAndGeneratedFilesFolderPath + @"test-reports-log.xlsx";
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

            ConsoleMessage($"AOI unit testing for '{testObjectName}' completed in {formattedTestLength} (HH:mm:ss).", "NEWSECTION");

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
            ConsoleMessage("START file management of content created during current and previous testing...", "NEWSECTION", false);
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

            // Rename text file to include the target object name.
            string textFileReportPath = reportAndGeneratedFilesFolderPath + @"1-textreports\" + currentDateTime + "_UnitTestReport.txt";
            RenameFile(textFileReportPath, currentDateTime + "_" + testObjectName + "_UnitTestReport.txt");
            #endregion
        }

        #region METHODS: L5X Manipulation
        /// <summary>
        /// Get the Name, DataType, Usage, Required, and Visible components of each parameter from an AOI definition XML file.
        /// </summary>
        /// <param name="l5xPath">The file path to the AOI definition L5X file.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>An array of the AOIParameter structure. Each array element corresponds to a unique AOI parameter.</returns>
        private static AOIParameter[]? GetParametersFromL5X(string l5xPath, bool printOut = false)
        {
            AOIParameter[]? returnAOIParams = null;
            int parameterCount;
            int paramIndex = 0;

            XDocument xDoc = XDocument.Load(l5xPath);
            var parameters = xDoc
                .Descendants("AddOnInstructionDefinition")
                .Descendants("Parameters")
                .Elements("Parameter")
                .Where(p => (string)p.Attribute("Usage")! != "InOut");
            parameterCount = parameters.Count();
            returnAOIParams = new AOIParameter[parameterCount];

            foreach (var p in parameters)
            {
                returnAOIParams[paramIndex].Name = p.Attribute("Name")!.Value;
                returnAOIParams[paramIndex].DataType = p.Attribute("DataType")!.Value;
                returnAOIParams[paramIndex].Usage = p.Attribute("Usage")!.Value;
                returnAOIParams[paramIndex].Required = bool.Parse(p.Attribute("Required")!.Value);
                returnAOIParams[paramIndex].Visible = bool.Parse(p.Attribute("Visible")!.Value)!;
                paramIndex++;
            }

            if (printOut)
                ConsoleMessage("Got complex tag sub componenet Name, Data Type, Usage, Required, & Visible values from L5X.", "STATUS");

            return returnAOIParams;
        }

        /// <summary>
        /// Convert an AOI definition L5X file to a routine L5X file with the following characteristics:<br/>
        ///  - The new routine contains 1 rung with an XIC instruction in series with a populated instance of the AOI instruction.<br/>
        ///  - The XIC instruction contains a newly created BOOL tag that is toggled depending on EnableIn logic.<br/>
        ///  - The AOI instruction has all required/visible parameters populated by a newly created AOI tag.
        /// </summary>
        /// <param name="l5xFilePath">The file path to the AOI definition L5X file.</param>
        /// <param name="routineName">The name of the new routine being created.</param>
        /// <param name="programName">The name of the program to which the routine is being added.</param>
        /// <param name="controllerName">The name of the controller to which the routine is being added.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void ConvertAOIL5X_DEFINITIONtoROUTINE(string l5xFilePath, string routineName, string programName, string controllerName, bool printOut)
        {
            string aoiName = GetAttributeValue(l5xFilePath, "AddOnInstructionDefinition", "Name", printOut)!;

            // Modify the "top" of the L5X file. This converts the definition L5X to a routine L5X.
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "TargetName", routineName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "TargetType", "Routine", printOut);
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "TargetSubType", "RLL", printOut);
            DeleteAttributeFromRoot(l5xFilePath, "TargetRevision", printOut);
            DeleteAttributeFromRoot(l5xFilePath, "TargetLastEdited", printOut);
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "ExportDate", DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy"), printOut);

            DeleteAttributeFromComplexElement(l5xFilePath, "AddOnInstructionDefinition", "Use", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Controller", "Name", controllerName, printOut);

            // Modify the "bottom" of the L5X file. This creates the routine, ladder logic rung, and tags as required in order to execute unit testing.
            AddElementToComplexElement(l5xFilePath, "Controller", "Tags", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tags", "Use", "Context", printOut);

            // Create the AOI tag instance using the existing XML information provided in the defintion L5X file.
            AddElementToComplexElement(l5xFilePath, "Tags", "Tag", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Name", "AOI_" + aoiName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "TagType", "Base", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "DataType", aoiName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Constant", "false", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "OpcUaAccess", "None", printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Description", printOut);
            string cdataforAOI_Descr = @"Automated Testing -------------------- generated AOI tag being unit tested";
            AddCDATA(l5xFilePath, "Description", cdataforAOI_Descr, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "L5K", printOut);
            string cdataInfo_forData = GetAOITagCDATA_forData(l5xFilePath, aoiName, printOut)!;
            AddCDATA(l5xFilePath, "Data", cdataInfo_forData, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "Decorated", printOut);
            AddElementToComplexElement(l5xFilePath, "Data", "Structure", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Structure", "DataType", aoiName, printOut);
            List<Dictionary<string, string>> attributesList = GetDataValueMemberInfofromXML(l5xFilePath, printOut);
            AddComplexElementsWithAttributesToXml(l5xFilePath, attributesList, printOut);

            // Add all InOut tags to the Tags complex element.
            try
            {
                XDocument doc = XDocument.Load(l5xFilePath);

                // Find all "Parameter" elements.
                var parameterElements = doc.Descendants("Parameter");

                // Cycle through each AOI parameter and add it to the list if it is a required parameter.
                foreach (var param in parameterElements)
                {
                    XAttribute? nameAttribute = param.Attribute("Name");
                    XAttribute? externalAccessAttribute = param.Attribute("ExternalAccess");
                    string requiredAttributeValue = param.Attribute("Required")!.Value;
                    string dataTypeAttributeValue = param.Attribute("DataType")!.Value;
                    // Create a new tag for each string InOut parameter.
                    if (nameAttribute != null && externalAccessAttribute == null && requiredAttributeValue == "true")
                    {
                        AddElementToComplexElement(l5xFilePath, "Tags", "Tag", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "Name", nameAttribute.Value, printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "TagType", "Base", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "DataType", dataTypeAttributeValue, printOut);
                        if (dataTypeAttributeValue == "REAL")
                            AddAttributeToComplexElement(l5xFilePath, "Tag", "Radix", "Float", printOut);
                        else if (dataTypeAttributeValue != "String")
                            AddAttributeToComplexElement(l5xFilePath, "Tag", "Radix", "Decimal", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "Constant", "false", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "OpcUaAccess", "None", printOut);
                        AddElementToComplexElement(l5xFilePath, "Tag", "Description", printOut);
                        string cdataforCurrentInOutParam_Descr = $"Automated Testing -------------------- generated InOut param for AOI_{aoiName} tag";
                        AddCDATA(l5xFilePath, "Description", cdataforCurrentInOutParam_Descr, printOut);
                        AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);

                        AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "L5K", printOut);
                        string cdataforCurrentInOutParam_forData = (string)GetInOutParamTagCDATA_forData(l5xFilePath, nameAttribute.Value)[0];
                        AddCDATA(l5xFilePath, "Data", cdataforCurrentInOutParam_forData, printOut);
                        AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
                        if (dataTypeAttributeValue != "String")
                        {
                            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "Decorated", printOut);
                            AddElementToComplexElement(l5xFilePath, "Data", "DataValue", printOut);
                            AddAttributeToComplexElement(l5xFilePath, "DataValue", "DataType", dataTypeAttributeValue, printOut);
                            if (dataTypeAttributeValue == "REAL")
                                AddAttributeToComplexElement(l5xFilePath, "DataValue", "Radix", "Float", printOut);
                            else
                                AddAttributeToComplexElement(l5xFilePath, "DataValue", "Radix", "Decimal", printOut);
                            AddAttributeToComplexElement(l5xFilePath, "DataValue", "Value", GetS5kAtomicTagCDATA_forData(dataTypeAttributeValue)!, printOut);

                        }
                        else
                        {
                            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "String", printOut);
                            AddAttributeToComplexElement(l5xFilePath, "Data", "Length", "0", printOut);
                            AddCDATA(l5xFilePath, "Data", "''", printOut);

                        }
                    }
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }

            // AT_EnableIn tag
            AddElementToComplexElement(l5xFilePath, "Tags", "Tag", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Name", "AT_EnableIn", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "TagType", "Base", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "DataType", "BOOL", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Radix", "Decimal", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Constant", "false", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "OpcUaAccess", "None", printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Description", printOut);
            string cdataforAT_EnableTag_Descr = @"Automated Testing -------------------- set the EnableIn AOI input parameter";
            AddCDATA(l5xFilePath, "Description", cdataforAT_EnableTag_Descr, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "L5K", printOut);
            string cdataforAT_EnableTag_Data = @"0";
            AddCDATA(l5xFilePath, "Data", cdataforAT_EnableTag_Data, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "Decorated", printOut);
            AddElementToComplexElement(l5xFilePath, "Data", "DataValue", printOut);
            AddAttributeToComplexElement(l5xFilePath, "DataValue", "DataType", "BOOL", printOut);
            AddAttributeToComplexElement(l5xFilePath, "DataValue", "Radix", "Decimal", printOut);
            AddAttributeToComplexElement(l5xFilePath, "DataValue", "Value", "0", printOut);

            AddElementToComplexElement(l5xFilePath, "Controller", "Programs", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Programs", "Use", "Context", printOut);

            AddElementToComplexElement(l5xFilePath, "Programs", "Program", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Program", "Use", "Context", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Program", "Name", programName, printOut);
            AddElementToComplexElement(l5xFilePath, "Program", "Routines", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routines", "Use", "Context", printOut);

            AddElementToComplexElement(l5xFilePath, "Routines", "Routine", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routine", "Use", "Target", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routine", "Name", routineName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routine", "Type", "RLL", printOut);

            AddElementToComplexElement(l5xFilePath, "Routine", "RLLContent", printOut);
            AddElementToComplexElement(l5xFilePath, "RLLContent", "Rung", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Rung", "Number", "0", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Rung", "Type", "N", printOut);

            AddElementToComplexElement(l5xFilePath, "Rung", "Comment", printOut);
            string cdataInfoforRung0Comment = @"AUTOMATED TESTING | " + aoiName + @" AOI UNIT TEST
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
This a programmatically created rung with a populated instance of the AOI instruction added using the Logix Designer SDK.";
            AddCDATA(l5xFilePath, "Comment", cdataInfoforRung0Comment, printOut);

            AddElementToComplexElement(l5xFilePath, "Rung", "Text", printOut);
            string cdataInfo_forText = GetCDATAfromXML_forText(l5xFilePath, printOut);
            AddCDATA(l5xFilePath, "Text", cdataInfo_forText, printOut);
        }

        /// <summary>
        /// Get the CDATA contents needed to create a Studio 5000 Logix Designer Tag for an AOI InOut parameter.
        /// </summary>
        /// <param name="l5xFilePath">The AOI definition L5X file path.</param>
        /// <param name="paramName">The name of the target parameter.</param>
        /// <returns>A string of CDATA contents needed to create a Studio 5000 Logix Designer tag.</returns>
        /// <exception cref="ArgumentException">Thrown if integers for tag dimension size cannot be parsed properly from L5X file.</exception>
        public static object[] GetInOutParamTagCDATA_forData(string l5xFilePath, string paramName, int boolCount = 0, bool printOut = false)
        {
            // The variable returnCDATA contains the final result to be returned.
            StringBuilder returnCDATA = new();
            object[] returnObjectArray = new object[2];

            // Variables required to account for tag nesting.
            StringBuilder nested0Contents = new();
            StringBuilder nested1Contents = new();
            StringBuilder nested2Contents = new();
            StringBuilder nested3Contents = new();
            StringBuilder nested4Contents = new();
            StringBuilder nested5Contents = new();
            StringBuilder nested6Contents = new();
            StringBuilder nested7Contents = new();
            StringBuilder nested8Contents = new();
            string nested0DataType = "";
            string nested1DataType = "";
            string nested2DataType = "";
            string nested3DataType = "";
            string nested4DataType = "";
            string nested5DataType = "";
            string nested6DataType = "";
            string nested7DataType = "";
            string nested8DataType = "";
            int nested0Dimensions = 0;
            int nested1Dimensions = 0;
            int nested2Dimensions = 0;
            int nested3Dimensions = 0;
            int nested4Dimensions = 0;
            int nested5Dimensions = 0;
            int nested6Dimensions = 0;
            int nested7Dimensions = 0;
            int nested8Dimensions = 0;

            XDocument xdoc = XDocument.Load(l5xFilePath);

            // Find all "Parameter" elements from the AOI definition L5X file.
            var nested0ParameterElements = xdoc.Descendants("Parameter");
            var nested0LocalTagElements = xdoc.Descendants("LocalTag");

            // Cycle through each AOI parameter and get the nested level 0 target datatype and dimensions.
            foreach (var n0pe in nested0ParameterElements)
            {
                if (paramName == n0pe.Attribute("Name")!.Value)
                {
                    nested0DataType = n0pe.Attribute("DataType")!.Value;
                    if (n0pe.Attribute("Dimensions") != null)
                    {
                        if (!int.TryParse(n0pe.Attribute("Dimensions")!.Value, out nested0Dimensions))
                            throw new ArgumentException("Parameter dimensions must be a valid integer.");
                    }
                    break;
                }
                else
                {
                    foreach (var n0lte in nested0LocalTagElements)
                    {
                        if (paramName == n0lte.Attribute("Name")!.Value)
                        {
                            nested0DataType = n0lte.Attribute("DataType")!.Value;
                            if (n0lte.Attribute("Dimensions") != null)
                            {
                                if (!int.TryParse(n0lte.Attribute("Dimensions")!.Value, out nested0Dimensions))
                                    throw new ArgumentException("Parameter dimensions must be a valid integer.");
                            }
                            break;
                        }
                    }
                }
            }

            string? nested0CDATA = GetS5kAtomicTagCDATA_forData(nested0DataType, nested0Dimensions > 0);

            // If atomic data not in an array.
            if (nested0CDATA != null && nested0Dimensions == 0)
            {
                if (nested0DataType == "BOOL" || nested0DataType == "BIT")
                {
                    boolCount++;
                    if ((boolCount & 31) == 1)
                        returnCDATA.Append(nested0CDATA);
                }
                else
                    returnCDATA.Append(nested0CDATA);
            }
            // If atomic data in an array.
            else if (nested0CDATA != null && nested0Dimensions > 0)
            {
                nested0Contents.Append(CreateArrayCDATA(nested0CDATA, nested0Dimensions));
                returnCDATA.Append(nested0Contents.ToString());
            }
            // If data is complex.
            else if (nested0CDATA == null)
            {
                #region NESTED LEVELS: 1 to 8
                // Get the non-hidden members of the target DataType complex element.
                var nested1DataTypeMemberElements = xdoc
                    .Descendants("DataTypes")
                    .Elements("DataType")
                    .FirstOrDefault(dt => dt.Attribute("Name")!.Value == nested0DataType)!
                    .Descendants("Members")
                    .Elements("Member")
                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                // Rotate through the level 1 nested member elements & add them to the CDATA stringbuilder.
                nested1Contents.Append('[');
                foreach (var n1dtme in nested1DataTypeMemberElements)
                {
                    // Get the data type of the current nested level and the dimension of the current member.
                    nested1DataType = n1dtme.Attribute("DataType")!.Value;
                    if (n1dtme.Attribute("Dimension") != null)
                    {
                        if (!int.TryParse(n1dtme.Attribute("Dimension")!.Value, out nested1Dimensions))
                            throw new ArgumentException("Tag dimensions (at 1 nested level) must be a valid integer.");
                    }

                    string? nested1CDATA = GetS5kAtomicTagCDATA_forData(nested1DataType, nested1Dimensions > 0);

                    if (nested1CDATA != null && nested1Dimensions == 0)
                    {
                        if (nested1DataType == "BOOL" || nested1DataType == "BIT")
                        {
                            boolCount++;
                            if ((boolCount & 31) == 1)
                                nested1Contents.Append(nested1CDATA);
                        }
                        else
                            nested1Contents.Append(nested1CDATA);
                    }
                    else if (nested1CDATA != null && nested1Dimensions > 0)
                    {
                        nested1Contents.Append(CreateArrayCDATA(nested1CDATA, nested1Dimensions));
                    }
                    else if (nested1CDATA == null)
                    {
                        #region NESTED LEVELS: 2 to 8
                        var nested2DataTypeMemberElements = xdoc
                            .Descendants("DataTypes")
                            .Elements("DataType")
                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested1DataType)!
                            .Descendants("Members")
                            .Elements("Member")
                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                        // Rotate through the level 2 nested member elements & add them to the CDATA stringbuilder.
                        nested2Contents.Append('[');
                        foreach (var n2dtme in nested2DataTypeMemberElements)
                        {
                            nested2DataType = n2dtme.Attribute("DataType")!.Value;
                            if (n2dtme.Attribute("Dimension") != null)
                            {
                                if (!int.TryParse(n2dtme.Attribute("Dimension")!.Value, out nested2Dimensions))
                                    throw new ArgumentException("Tag dimensions (at 2 nested level) must be a valid integer.");
                            }

                            string? nested2CDATA = GetS5kAtomicTagCDATA_forData(nested2DataType, nested2Dimensions > 0);

                            if (nested2CDATA != null && nested2Dimensions == 0)
                            {
                                if (nested2DataType == "BOOL" || nested2DataType == "BIT")
                                {
                                    boolCount++;
                                    if ((boolCount & 31) == 1)
                                        nested2Contents.Append(nested2CDATA);
                                }
                                else
                                    nested2Contents.Append(nested2CDATA);
                            }
                            else if (nested2CDATA != null && nested2Dimensions > 0)
                            {
                                nested2Contents.Append(CreateArrayCDATA(nested2CDATA, nested2Dimensions));
                            }
                            else if (nested2CDATA == null)
                            {
                                #region NESTED LEVELS: 3 to 8
                                var nested3DataTypeMemberElements = xdoc
                                    .Descendants("DataTypes")
                                    .Elements("DataType")
                                    .FirstOrDefault(e => (string)e.Attribute("Name")! == nested2DataType)!
                                    .Descendants("Members")
                                    .Elements("Member")
                                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                                // Rotate through the level 3 nested member elements & add them to the CDATA stringbuilder.
                                nested3Contents.Append('[');
                                foreach (var n3dtme in nested3DataTypeMemberElements)
                                {
                                    nested3DataType = n3dtme.Attribute("DataType")!.Value;
                                    if (n3dtme.Attribute("Dimension") != null)
                                    {
                                        if (!int.TryParse(n3dtme.Attribute("Dimension")!.Value, out nested3Dimensions))
                                            throw new ArgumentException("Tag dimensions (at 3 nested level) must be a valid integer.");
                                    }

                                    string? nested3CDATA = GetS5kAtomicTagCDATA_forData(nested3DataType, nested3Dimensions > 0);

                                    if (nested3CDATA != null && nested3Dimensions == 0)
                                    {
                                        if (nested3DataType == "BOOL" || nested3DataType == "BIT")
                                        {
                                            boolCount++;
                                            if ((boolCount & 31) == 1)
                                                nested3Contents.Append(nested3CDATA);
                                        }
                                        else
                                            nested3Contents.Append(nested3CDATA);
                                    }
                                    else if (nested3CDATA != null && nested3Dimensions > 0)
                                    {
                                        nested3Contents.Append(CreateArrayCDATA(nested3CDATA, nested3Dimensions));
                                    }
                                    else if (nested3CDATA == null)
                                    {
                                        #region NESTED LEVELS: 4 to 8
                                        var nested4DataTypeMemberElements = xdoc
                                            .Descendants("DataTypes")
                                            .Elements("DataType")
                                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested3DataType)!
                                            .Descendants("Members")
                                            .Elements("Member")
                                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                                        // Rotate through the level 4 nested member elements & add them to the CDATA stringbuilder.
                                        nested4Contents.Append('[');
                                        foreach (var n4dtme in nested4DataTypeMemberElements)
                                        {
                                            nested4DataType = n4dtme.Attribute("DataType")!.Value;
                                            if (n4dtme.Attribute("Dimension") != null)
                                            {
                                                if (!int.TryParse(n4dtme.Attribute("Dimension")!.Value, out nested4Dimensions))
                                                    throw new ArgumentException("Tag dimensions (at 4 nested level) must be a valid integer.");
                                            }

                                            string? nested4CDATA = GetS5kAtomicTagCDATA_forData(nested4DataType, nested4Dimensions > 0);

                                            if (nested4CDATA != null && nested4Dimensions == 0)
                                            {
                                                if (nested4DataType == "BOOL" || nested4DataType == "BIT")
                                                {
                                                    boolCount++;
                                                    if ((boolCount & 31) == 1)
                                                        nested4Contents.Append(nested4CDATA);
                                                }
                                                else
                                                    nested4Contents.Append(nested4CDATA);
                                            }
                                            else if (nested4CDATA != null && nested4Dimensions > 0)
                                            {
                                                nested4Contents.Append(CreateArrayCDATA(nested4CDATA, nested4Dimensions));
                                            }
                                            else if (nested4CDATA == null)
                                            {
                                                #region NESTED LEVELS: 5 to 8
                                                var nested5DataTypeMemberElements = xdoc
                                                    .Descendants("DataTypes")
                                                    .Elements("DataType")
                                                    .FirstOrDefault(e => (string)e.Attribute("Name")! == nested4DataType)!
                                                    .Descendants("Members")
                                                    .Elements("Member")
                                                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                // Rotate through the level 5 nested member elements & add them to the CDATA stringbuilder.
                                                nested5Contents.Append('[');
                                                foreach (var n5dtme in nested5DataTypeMemberElements)
                                                {
                                                    nested5DataType = n5dtme.Attribute("DataType")!.Value;
                                                    if (n5dtme.Attribute("Dimension") != null)
                                                    {
                                                        if (!int.TryParse(n5dtme.Attribute("Dimension")!.Value, out nested5Dimensions))
                                                            throw new ArgumentException("Tag dimensions (at 5 nested level) must be a valid integer.");
                                                    }

                                                    string? nested5CDATA = GetS5kAtomicTagCDATA_forData(nested5DataType, nested5Dimensions > 0);

                                                    if (nested5CDATA != null && nested5Dimensions == 0)
                                                    {
                                                        if (nested5DataType == "BOOL" || nested5DataType == "BIT")
                                                        {
                                                            boolCount++;
                                                            if ((boolCount & 31) == 1)
                                                                nested5Contents.Append(nested5CDATA);
                                                        }
                                                        else
                                                            nested5Contents.Append(nested5CDATA);
                                                    }
                                                    else if (nested5CDATA != null && nested5Dimensions > 0)
                                                    {
                                                        nested5Contents.Append(CreateArrayCDATA(nested5CDATA, nested5Dimensions));
                                                    }
                                                    else if (nested5CDATA == null)
                                                    {
                                                        #region NESTED LEVELS: 6 to 8
                                                        var nested6DataTypeMemberElements = xdoc
                                                            .Descendants("DataTypes")
                                                            .Elements("DataType")
                                                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested5DataType)!
                                                            .Descendants("Members")
                                                            .Elements("Member")
                                                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                        // Rotate through the level 6 nested member elements & add them to the CDATA stringbuilder.
                                                        nested6Contents.Append('[');
                                                        foreach (var n6dtme in nested6DataTypeMemberElements)
                                                        {
                                                            nested6DataType = n6dtme.Attribute("DataType")!.Value;
                                                            if (n6dtme.Attribute("Dimension") != null)
                                                            {
                                                                if (!int.TryParse(n6dtme.Attribute("Dimension")!.Value, out nested6Dimensions))
                                                                    throw new ArgumentException("Tag dimensions (at 6 nested level) must be a valid integer.");
                                                            }

                                                            string? nested6CDATA = GetS5kAtomicTagCDATA_forData(nested6DataType, nested6Dimensions > 0);

                                                            if (nested6CDATA != null && nested6Dimensions == 0)
                                                            {
                                                                if (nested6DataType == "BOOL" || nested6DataType == "BIT")
                                                                {
                                                                    boolCount++;
                                                                    if ((boolCount & 31) == 1)
                                                                        nested6Contents.Append(nested6CDATA);
                                                                }
                                                                else
                                                                    nested6Contents.Append(nested6CDATA);
                                                            }
                                                            else if (nested6CDATA != null && nested6Dimensions > 0)
                                                            {
                                                                nested6Contents.Append(CreateArrayCDATA(nested6CDATA, nested6Dimensions));
                                                            }
                                                            else if (nested6CDATA == null)
                                                            {
                                                                #region NESTED LEVELS: 7 to 8
                                                                var nested7DataTypeMemberElements = xdoc
                                                                    .Descendants("DataTypes")
                                                                    .Elements("DataType")
                                                                    .FirstOrDefault(e => (string)e.Attribute("Name")! == nested6DataType)!
                                                                    .Descendants("Members")
                                                                    .Elements("Member")
                                                                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                                // Rotate through the level 7 nested member elements & add them to the CDATA stringbuilder.
                                                                nested7Contents.Append('[');
                                                                foreach (var n7dtme in nested7DataTypeMemberElements)
                                                                {
                                                                    nested7DataType = n7dtme.Attribute("DataType")!.Value;
                                                                    if (n7dtme.Attribute("Dimension") != null)
                                                                    {
                                                                        if (!int.TryParse(n7dtme.Attribute("Dimension")!.Value, out nested7Dimensions))
                                                                            throw new ArgumentException("Tag dimensions (at 7 nested level) must be a valid integer.");
                                                                    }

                                                                    string? nested7CDATA = GetS5kAtomicTagCDATA_forData(nested7DataType, nested7Dimensions > 0);

                                                                    if (nested7CDATA != null && nested7Dimensions == 0)
                                                                    {
                                                                        if (nested7DataType == "BOOL" || nested7DataType == "BIT")
                                                                        {
                                                                            boolCount++;
                                                                            if ((boolCount & 31) == 1)
                                                                                nested7Contents.Append(nested7CDATA);
                                                                        }
                                                                        else
                                                                            nested7Contents.Append(nested7CDATA);
                                                                    }
                                                                    else if (nested7CDATA != null && nested7Dimensions > 0)
                                                                    {
                                                                        nested7Contents.Append(CreateArrayCDATA(nested7CDATA, nested7Dimensions));
                                                                    }
                                                                    else if (nested7CDATA == null)
                                                                    {
                                                                        #region NESTED LEVELS: 8 to 8
                                                                        var nested8DataTypeMemberElements = xdoc
                                                                            .Descendants("DataTypes")
                                                                            .Elements("DataType")
                                                                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested7DataType)!
                                                                            .Descendants("Members")
                                                                            .Elements("Member")
                                                                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                                        // Rotate through the level 8 nested member elements & add them to the CDATA stringbuilder.
                                                                        nested8Contents.Append('[');
                                                                        foreach (var n8dtme in nested8DataTypeMemberElements)
                                                                        {
                                                                            nested8DataType = n8dtme.Attribute("DataType")!.Value;
                                                                            if (n8dtme.Attribute("Dimension") != null)
                                                                            {
                                                                                if (!int.TryParse(n8dtme.Attribute("Dimension")!.Value, out nested8Dimensions))
                                                                                    throw new ArgumentException("Tag dimensions (at 8 nested level) must be a valid integer.");
                                                                            }

                                                                            string? nested8CDATA = GetS5kAtomicTagCDATA_forData(nested8DataType, nested8Dimensions > 0);

                                                                            if (nested8CDATA != null && nested8Dimensions == 0)
                                                                            {
                                                                                if (nested8DataType == "BOOL" || nested8DataType == "BIT")
                                                                                {
                                                                                    boolCount++;
                                                                                    if ((boolCount & 31) == 1)
                                                                                        nested8Contents.Append(nested8CDATA);
                                                                                }
                                                                                else
                                                                                    nested8Contents.Append(nested8CDATA);
                                                                            }
                                                                            else if (nested8CDATA != null && nested8Dimensions > 0)
                                                                            {
                                                                                nested8Contents.Append(CreateArrayCDATA(nested8CDATA, nested8Dimensions));
                                                                            }
                                                                            else if (nested8CDATA == null)
                                                                            {
                                                                                ConsoleMessage("Nesting tags beyond 8 levels not supported.", "ERROR");
                                                                            }

                                                                            // Add a comma for all situations other than boolean "bit packing".
                                                                            if ((boolCount & 31) == 1)
                                                                                nested8Contents.Append(',');
                                                                            else if (nested8DataType != "BIT")
                                                                                nested8Contents.Append(',');
                                                                        }
                                                                        nested8Contents.Length--;    // Remove the last apostrophe character ','
                                                                        nested8Contents.Append(']'); // Finish the nested component syntax.

                                                                        if (nested7Dimensions == 0)
                                                                        {
                                                                            nested7Contents.Append(nested8Contents.ToString());
                                                                        }
                                                                        else if (nested7Dimensions > 0)
                                                                        {
                                                                            nested7Contents.Append(CreateArrayCDATA(nested8Contents.ToString(), nested7Dimensions));
                                                                        }

                                                                        nested8Contents.Clear();
                                                                        #endregion
                                                                    }

                                                                    // Add a comma for all situations other than boolean "bit packing".
                                                                    if ((boolCount & 31) == 1)
                                                                        nested7Contents.Append(',');
                                                                    else if (nested7DataType != "BIT")
                                                                        nested7Contents.Append(',');
                                                                }
                                                                nested7Contents.Length--;    // Remove the last apostrophe character ','
                                                                nested7Contents.Append(']'); // Finish the nested component syntax.

                                                                if (nested6Dimensions == 0)
                                                                {
                                                                    nested6Contents.Append(nested7Contents.ToString());
                                                                }
                                                                else if (nested6Dimensions > 0)
                                                                {
                                                                    nested6Contents.Append(CreateArrayCDATA(nested7Contents.ToString(), nested6Dimensions));
                                                                }

                                                                nested7Contents.Clear();
                                                                #endregion
                                                            }

                                                            // Add a comma for all situations other than boolean "bit packing".
                                                            if ((boolCount & 31) == 1)
                                                                nested6Contents.Append(',');
                                                            else if (nested6DataType != "BIT")
                                                                nested6Contents.Append(',');
                                                        }
                                                        nested6Contents.Length--;    // Remove the last apostrophe character ','
                                                        nested6Contents.Append(']'); // Finish the nested component syntax.

                                                        if (nested5Dimensions == 0)
                                                        {
                                                            nested5Contents.Append(nested6Contents.ToString());
                                                        }
                                                        else if (nested5Dimensions > 0)
                                                        {
                                                            nested5Contents.Append(CreateArrayCDATA(nested6Contents.ToString(), nested5Dimensions));
                                                        }

                                                        nested6Contents.Clear();
                                                        #endregion
                                                    }

                                                    // Add a comma for all situations other than boolean "bit packing".
                                                    if ((boolCount & 31) == 1)
                                                        nested5Contents.Append(',');
                                                    else if (nested5DataType != "BIT")
                                                        nested5Contents.Append(',');
                                                }
                                                nested5Contents.Length--;    // Remove the last apostrophe character ','
                                                nested5Contents.Append(']'); // Finish the nested component syntax.

                                                if (nested4Dimensions == 0)
                                                {
                                                    nested4Contents.Append(nested5Contents.ToString());
                                                }
                                                else if (nested4Dimensions > 0)
                                                {
                                                    nested4Contents.Append(CreateArrayCDATA(nested5Contents.ToString(), nested4Dimensions));
                                                }

                                                nested5Contents.Clear();
                                                #endregion
                                            }

                                            // Add a comma for all situations other than boolean "bit packing".
                                            if ((boolCount & 31) == 1)
                                                nested4Contents.Append(',');
                                            else if (nested4DataType != "BIT")
                                                nested4Contents.Append(',');
                                        }
                                        nested4Contents.Length--;    // Remove the last apostrophe character ','
                                        nested4Contents.Append(']'); // Finish the nested component syntax.

                                        if (nested3Dimensions == 0)
                                        {
                                            nested3Contents.Append(nested4Contents.ToString());
                                        }
                                        else if (nested3Dimensions > 0)
                                        {
                                            nested3Contents.Append(CreateArrayCDATA(nested4Contents.ToString(), nested3Dimensions));
                                        }

                                        nested4Contents.Clear();
                                        #endregion

                                    }

                                    // Add a comma for all situations other than boolean "bit packing".
                                    if ((boolCount & 31) == 1)
                                        nested3Contents.Append(',');
                                    else if (nested3DataType != "BIT")
                                        nested3Contents.Append(',');
                                }
                                nested3Contents.Length--;    // Remove the last apostrophe character ','
                                nested3Contents.Append(']'); // Finish the nested component syntax.

                                if (nested2Dimensions == 0)
                                {
                                    nested2Contents.Append(nested3Contents.ToString());
                                }
                                else if (nested2Dimensions > 0)
                                {
                                    nested2Contents.Append(CreateArrayCDATA(nested3Contents.ToString(), nested2Dimensions));
                                }

                                nested3Contents.Clear();
                                #endregion
                            }

                            // Add a comma for all situations other than boolean "bit packing".
                            if ((boolCount & 31) == 1)
                                nested2Contents.Append(',');
                            else if (nested2DataType != "BIT")
                                nested2Contents.Append(',');
                        }
                        nested2Contents.Length--;    // Remove the last apostrophe character ','
                        nested2Contents.Append(']'); // Finish the nested component syntax.

                        if (nested1Dimensions == 0)
                        {
                            nested1Contents.Append(nested2Contents.ToString());
                        }
                        else if (nested1Dimensions > 0)
                        {
                            nested1Contents.Append(CreateArrayCDATA(nested2Contents.ToString(), nested1Dimensions));
                        }

                        nested2Contents.Clear();
                        #endregion
                    }

                    // Add a comma for all situations other than boolean "bit packing".
                    if ((boolCount & 31) == 1)
                        nested1Contents.Append(',');
                    else if (nested1DataType != "BIT")
                        nested1Contents.Append(',');
                }
                nested1Contents.Length--;    // Remove the last apostrophe character ','
                nested1Contents.Append(']'); // Finish the nested component syntax.

                if (nested0Dimensions == 0)
                {
                    returnCDATA.Append(nested1Contents.ToString());
                }
                else if (nested0Dimensions > 0)
                {
                    returnCDATA.Append(CreateArrayCDATA(nested1Contents.ToString(), nested0Dimensions));
                }

                nested1Contents.Clear();
                #endregion
            }

            if (printOut)
                ConsoleMessage($"CDATA contents: {returnCDATA.ToString()}", "STATUS");

            returnObjectArray[0] = returnCDATA.ToString();
            returnObjectArray[1] = boolCount;
            return returnObjectArray;
        }

        /// <summary>
        /// Programmatically get the CDATA contents for the 'Data' complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A string of formatted CDATA contents.</returns>
        public static string? GetAOITagCDATA_forData(string xmlFilePath, string aoiName, bool printOut)
        {
            int boolCount = 0;
            List<string> paramElemStringList = new List<string>();
            List<string> localTagElemStringList = new List<string>();

            XDocument doc = XDocument.Load(xmlFilePath);

            // Get a list filtered to contain only CDATA information from nonboolean "Parameter" elements.
            var parameterElements = doc
                .Descendants("AddOnInstructionDefinition")
                .Where(param => param.Attribute("Name")?.Value == aoiName)
                .Descendants("Parameters")
                .Elements("Parameter")
                .Where(param => param.Attribute("Usage")?.Value != "InOut");

            foreach (var pe in parameterElements)
            {
                object[] currentParamCDATA = GetInOutParamTagCDATA_forData(xmlFilePath, pe.Attribute("Name")!.Value, boolCount);
                string currentCDATA = (string)currentParamCDATA[0];
                boolCount = 0;
                boolCount += (int)currentParamCDATA[1];

                if (currentCDATA != "")
                    paramElemStringList.Add(currentCDATA);
            }

            // Join all parameterElements list elements into a single string, with each element separated by a comma without spaces.
            string joined_pCDATA = string.Join(",", paramElemStringList);

            // Get a list filtered to contain only CDATA information from nonboolean "LocalTag" elements.
            var localtagElements = doc
                .Descendants("AddOnInstructionDefinition")
                .Where(param => param.Attribute("Name")?.Value == aoiName)
                .Descendants("LocalTags")
                .Elements("LocalTag")
                .Where(param => param.Attribute("Usage")?.Value != "InOut");

            foreach (var lte in localtagElements)
            {
                object[] currentLocalTagCDATA = GetInOutParamTagCDATA_forData(xmlFilePath, lte.Attribute("Name")!.Value, boolCount);
                string currentCDATA = (string)currentLocalTagCDATA[0];
                boolCount = 0;
                boolCount += (int)currentLocalTagCDATA[1];

                if (currentCDATA != "")
                    localTagElemStringList.Add(currentCDATA);
            }

            // Join all localtagElements list elements into a single string, with each element separated by a comma without spaces.
            string joined_ltCDATA = string.Join(",", localTagElemStringList);

            // If no local tags, don't include the joined local tag CDATA information.
            string returnString;
            if (joined_ltCDATA == "")
            {
                returnString = "[" + joined_pCDATA + "]";
            }
            else
            {
                returnString = "[" + joined_pCDATA + "," + joined_ltCDATA + "]";
            }

            // Create the final formatted string to be used as CDATA content information (in the Data complex element of L5X).
            if (printOut)
                ConsoleMessage($"CDATA contents: {returnString}", "STATUS");

            return returnString;
        }

        /// <summary>
        /// Get the value of an attribute for a specific complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
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

        /// <summary>
        /// Delete an attribute (name & value) from a complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element containing the attribute to be deleted.</param>
        /// <param name="attributeToDelete">The name of the attribute to be deleted.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void DeleteAttributeFromComplexElement(string xmlFilePath, string complexElementName, string attributeToDelete, bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).FirstOrDefault()!;

                if (complexElement != null)
                {
                    XAttribute attribute = complexElement.Attribute(attributeToDelete)!;

                    if (attribute != null)
                    {
                        attribute.Remove();

                        if (printOut)
                        {
                            ConsoleMessage($"Attribute '{attributeToDelete}' has been removed from the element '{complexElementName}'.", "STATUS");
                        }

                        xdoc.Save(xmlFilePath);
                    }
                    else if (printOut)
                    {
                        ConsoleMessage($"Attribute '{attributeToDelete}' not found in element '{complexElementName}'.", "ERROR");
                    }
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Delete an attribute from the root complex element of a 
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="attributeToDelete">The name of the attribute to be deleted.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        public static void DeleteAttributeFromRoot(string xmlFilePath, string attributeToDelete, bool printOut)
        {
            XDocument xdoc = XDocument.Load(xmlFilePath);
            XElement root = xdoc.Root!;
            XAttribute attribute = root.Attribute(attributeToDelete)!;

            if (attribute != null)
            {
                attribute.Remove();

                if (printOut)
                {
                    string complexElementName = "RSLogix5000Content";
                    ConsoleMessage($"Attribute '{attributeToDelete}' has been removed from the root complex element '{complexElementName}'.", "STATUS");
                }

                xdoc.Save(xmlFilePath);
            }
            else if (printOut)
            {
                ConsoleMessage($"Attribute '{attributeToDelete}' not found in the root complex element.", "ERROR");
            }
        }

        /// <summary>
        /// Add (or overwrite) an attribute name and value to a complex element in XML.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the attribute will be added.</param>
        /// <param name="attributeName">The name of the attribute to be added.</param>
        /// <param name="attributeValue">The value of the attirbute to be added.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        public static void AddAttributeToComplexElement(string xmlFilePath, string complexElementName, string attributeName, string attributeValue,
            bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault()!;

                if (complexElement != null)
                {
                    complexElement.SetAttributeValue(attributeName, attributeValue);

                    if (printOut)
                    {
                        ConsoleMessage($"Attribute '{attributeName}' with value '{attributeValue}' has been added to the element " +
                            $"'{complexElementName}'.", "STATUS");
                    }

                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Add an element to a complex element in XML.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the element will be added.</param>
        /// <param name="newElementName">The name of the new element.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        public static void AddElementToComplexElement(string xmlFilePath, string complexElementName, string newElementName, bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault()!;

                if (complexElement != null)
                {
                    XElement newElement = new XElement(newElementName);
                    complexElement.Add(newElement);

                    if (printOut)
                    {
                        ConsoleMessage($"Element '{newElementName}' has been added to the complex element '{complexElementName}'.", "STATUS");
                    }

                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Create a new CDATA element to the last or default instance of a specified complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the CDATA element will be added.</param>
        /// <param name="cdataContent">The contents of the CDATA element.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void AddCDATA(string xmlFilePath, string complexElementName, string cdataContent, bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault()!;

                if (complexElement != null)
                {
                    XCData cdataSection = new XCData(cdataContent);
                    complexElement.Add(cdataSection);

                    if (printOut)
                    {
                        ConsoleMessage($"A new CDATA section has been created and added to the element '{complexElementName}'.", "STATUS");
                    }

                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Programmatically get the CDATA contents for the Text complex element.<br/>
        /// This method programmatically gathers and formats where the information needed for a new AOI tag.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A string of formatted CDATA contents.</returns>
        public static string GetCDATAfromXML_forText(string xmlFilePath, bool printOut)
        {
            // The name of the AOI being tested.
            string? aoiName = GetAttributeValue(xmlFilePath, "AddOnInstructionDefinition", "Name", printOut);

            // Initialize the StringBuilder that will contain the AOI parameter tag names.
            StringBuilder aoiTagParameterNames = new();

            try
            {
                XDocument doc = XDocument.Load(xmlFilePath);

                // Find all "Parameter" elements.
                var parameterElements = doc.Descendants("Parameter");

                // Cycle through each AOI parameter and add it to the list if it is a required parameter.
                foreach (var param in parameterElements)
                {
                    XAttribute? nameAttribute = param.Attribute("Name");
                    XAttribute? externalAccessAttribute = param.Attribute("ExternalAccess");
                    string requiredAttributeValue = param.Attribute("Required")!.Value;

                    // Add all required parameters to the string that will populate the AOI instruction instance.
                    if (nameAttribute != null && requiredAttributeValue == "true")
                    {
                        if (externalAccessAttribute != null) // In/Out Params: no InOut because they do not have the ExternalAccess attribute.
                        {
                            aoiTagParameterNames.Append($",AOI_{aoiName}.{nameAttribute.Value}");
                        }
                        else // InOut parameters formatted such that they're not part of the AOI tag.
                        {
                            aoiTagParameterNames.Append("," + nameAttribute.Value);
                        }
                    }
                }

                string returnString = $"XIC(AT_EnableIn){aoiName}(AOI_{aoiName}{aoiTagParameterNames});";

                if (printOut)
                {
                    ConsoleMessage($"CDATA contents: {returnString}", "STATUS");
                }

                return returnString;
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
                return e.Message;
            }
        }

        /// <summary>
        /// Helper method for the GetInOutParamCDATA_forData method.<br/>
        /// Return the CDATA format of the specified datatype.
        /// </summary>
        /// <param name="dataType">The name of the target datatype.</param>
        /// <param name="IsArray">Boolean defaulted to false that varies what kind of CDATA the bool data type returns.</param>
        /// <returns>A string in the L5X CDATA format of a particular datatype.</returns>
        public static string? GetS5kAtomicTagCDATA_forData(string dataType, bool IsArray = false)
        {
            dataType = dataType.ToUpper();
            string boolcdataNoArray = "0";
            string boolcdataWithArray = "2#0";
            string sintcdata = "0";
            string intcdata = "0";
            string dintcdata = "0";
            string lintcdata = "0";
            string realcdata = "0.00000000e+000";
            string stringcdata = "[0,'$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00"
                          + "$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00']";
            string timercdata = "[0,0,0]";

            if (dataType == "BOOL" || dataType == "BIT")
            {
                if (IsArray)
                    return boolcdataWithArray;
                else
                    return boolcdataNoArray;
            }
            else if (dataType == "SINT")
            {
                return sintcdata;
            }
            else if (dataType == "INT")
            {
                return intcdata;
            }
            else if (dataType == "DINT")
            {
                return dintcdata;
            }
            else if (dataType == "LINT")
            {
                return lintcdata;
            }
            else if (dataType == "REAL")
            {
                return realcdata;
            }
            else if (dataType == "STRING")
            {
                return stringcdata;
            }
            else if (dataType == "TIMER")
            {
                return timercdata;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Helper method for the GetInOutParamCDATA_forData method.<br/>
        /// Create a the format needed for a CDATA array.
        /// </summary>
        /// <param name="objectToRepeat">The string to be repeated.</param>
        /// <param name="repeatCount">The number of times to repeat the string.</param>
        /// <returns>A formatted string in the L5X array format.</returns>
        public static string CreateArrayCDATA(string objectToRepeat, int repeatCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append('[');

            for (int i = 0; i < repeatCount; i++)
            {
                if (i > 0)
                {
                    sb.Append(',');
                }
                sb.Append(objectToRepeat);
            }

            sb.Append(']');
            return sb.ToString();
        }

        /// <summary>
        /// Get all the attribute names and values for each parameter in an AOI L5X file.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A list of dictionaries for each AOI parameter's attributes.</returns>
        public static List<Dictionary<string, string>> GetDataValueMemberInfofromXML(string xmlFilePath, bool printOut)
        {
            List<Dictionary<string, string>> return_attributeList = new List<Dictionary<string, string>>();

            try
            {
                XDocument doc = XDocument.Load(xmlFilePath);

                // Cycle through each "Parameter" element in the L5X file.
                foreach (var parameterElement in doc.Descendants("Parameter"))
                {
                    if (parameterElement.Attribute("Radix") != null && parameterElement.Attribute("Usage")!.Value != "InOut")
                    {
                        Dictionary<string, string> attributes = new Dictionary<string, string>
                        {
                            { "Name", parameterElement.Attribute("Name")!.Value },
                            { "DataType", parameterElement.Attribute("DataType")!.Value },
                            { "Radix", parameterElement.Attribute("Radix")!.Value }
                        };

                        // Store the new dictionary containing attributes for a single AOI parameter.
                        return_attributeList.Add(attributes);
                    }
                }

                if (printOut)
                    ConsoleMessage($"Got element information from L5X (Name, Data Type, & Radix) per AOI parameter.", "STATUS");
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }

            return return_attributeList;
        }

        /// <summary>
        /// For each AOI parameter, add the element "DataValueMember" with its attributes to the L5X complex element "Structure".<br/>
        /// This method creates XML children needed to create an AOI tag in the L5X file.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="attributesList">A list of dictionaries for each AOI parameter's attributes.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void AddComplexElementsWithAttributesToXml(string xmlFilePath, List<Dictionary<string, string>> attributesList, bool printOut)
        {
            try
            {
                foreach (var attributes in attributesList)
                {
                    // Add a new element "DataValueMember" to complex element "Structure" for each AOI parameter.
                    AddElementToComplexElement(xmlFilePath, "Structure", "DataValueMember", printOut);

                    // Add the "Name" attribute and its value for the current AOI parameter.
                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Name", attributes["Name"], printOut);

                    // Add the "DataType" attribute and its value for the current AOI parameter.
                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "DataType", attributes["DataType"], printOut);

                    // Add the "Radix" attribute and its value for the current AOI parameter.
                    // Note: BOOL datatype parameters don't have a "Radix" attribute and are therefore skipped.
                    if (attributes["DataType"] != "BOOL")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Radix", attributes["Radix"], printOut);
                    }

                    // Add the "Value" attribute and its value for the current AOI parameter.
                    // Note: For AOIs, the only BOOL parameter with a value of 1 is "EnableIn".
                    // Note: For REAL datatype parameters, their intial zero value has the notation "0.0". All else is "0".
                    if (attributes["Name"] == "EnableIn")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "1", printOut);
                    }
                    else if (attributes["DataType"] == "REAL")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0.0", printOut);
                    }
                    else
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0", printOut);
                    }
                }
                if (printOut)
                {
                    ConsoleMessage("Complex elements added.", "STATUS");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }
        #endregion

        #region METHODS: Excel Manipulation
        /// <summary>
        /// In the first worksheet of an Excel workbook, get the number of populated cells in the specified row.
        /// </summary>
        /// <param name="excelFilePath">The excel workbook file path.</param>
        /// <param name="rowNumber">The column in which the populated cell count is derived.</param>
        /// <returns>The number of populated cells in the specified row.</returns>
        private static int GetPopulatedCellsInRowCount(string excelFilePath, int rowNumber)
        {
            int returnCellCount = 0;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int maxColumnNum = worksheet.Dimension.End.Column;

                for (int col = 1; col <= maxColumnNum; col++)
                {
                    var cellValue = worksheet.Cells[rowNumber, col].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        returnCellCount++;
                }
            }
            return returnCellCount;
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
        /// Collect the values of each AOI parameter to be used during a test case from a specifically formatted excel workbook.
        /// </summary>
        /// <param name="excelFilePath">The file path of the excel workbook containing the test case information.</param>
        /// <param name="columnNumber">The column number of a test case in the excel file.</param>
        /// <returns>A dictionary where the Key is an AOI parameter name and the Value is an AOI parameter value.</returns>
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
        /// Create a copy of a specified file at a specific location for an AOI definition to routine conversion.
        /// </summary>
        /// <param name="sourceFilePath">The file to be copied.</param>
        /// <param name="outputFolderPath">The folder path in which the new file is to be saved.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>The file path of the copied file.</returns>
        public static string CopyFile(string sourceFilePath, string outputFolderPath, bool printOut = false)
        {
            if (!File.Exists(sourceFilePath))
                ConsoleMessage($"Source file '{sourceFilePath}' does not exist.", "ERROR");

            string targetObjectName = GetAttributeValue(sourceFilePath, "AddOnInstructionDefinition", "Name", printOut)!;

            // Get the directory and file name from the source file path.
            string extension = Path.GetExtension(sourceFilePath);

            // Construct the new file path for the copied file.
            string newFileName = $"{currentDateTime}_{targetObjectName}_AOIRoutine{extension}";
            string newFilePath = Path.Combine(outputFolderPath, newFileName);

            File.Copy(sourceFilePath, newFilePath, overwrite: true);

            return newFilePath;
        }
        #endregion
    }
}