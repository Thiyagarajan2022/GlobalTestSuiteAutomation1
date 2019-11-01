//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT EnvParams

var unitName, testCase, execute, description;
var TestCase_ID, functions, Execute, Data;

function executeTestCases(){
var instance = EnvParams.getEnvironment();
var businessFlow = EnvParams.getBusinessFlow();
var ReportDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y_%H.%M.%S");
Log.Message("HTML REPORT PATH::"+Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate);
Log.Message("-----------------------------------------------------");
ReportUtils.createConsolidatedReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", "ConsolidatedReport");
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
var rowcount = ExcelUtils.getRowCount()-1;
var excelRow=1;
while(excelRow<=rowcount){
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
unitName = ExcelUtils.getRowValue("UnitName",excelRow);
//Log.Message(unitName)
testCase = ExcelUtils.getRowValue("TestCases",excelRow);
//Log.Message(testCase)
description = ExcelUtils.getRowValue("Description",excelRow);
//Log.Message(description)
execute = ExcelUtils.getRowValue("Execute",excelRow);
//Log.Message(execute)
 
if(execute.toUpperCase()=="YES"){
reportName = "Report_"+unitName+"_"+testCase+".html"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest(unitName+" "+testCase, description)
Runner.CallMethod(unitName+"."+testCase);
// Runner.CallMethod("Caller.TestFlow",unitName,testCase);
}
// 
//    ExcelUtils.setExcelName(Project.Path+"\\DataTables\\TC_TestData.xlsx","General_Data");
//    var TestCaserowcount = ExcelUtils.getRowCount()-1;
//    var j=1;
//    while(j<=TestCaserowcount){
//
//     TestCase_ID = ExcelUtils.getRowValue("TestCase_ID",i);
//     functions = ExcelUtils.getRowValue("functions",i);
//     ExecuteTests = StrToInt(ExcelUtils.getRowValue("Execute",i));
//     TestIteration = ExcelUtils.getRowValue("Iteration",i);
//     Data = ExcelUtils.getRowValue("Data",i);
//     
//     if(ExecuteTests.toUpperCase()=="YES"){
//       
//     
//     
//     
//     }
//     
//     j++
//     }
// 
// }
 
excelRow++;
} 
//Aliases.browser.Close();

}

function asaa()
{
  Log.Message(TextUtils.GetValue("RunManagerPath"));
}

