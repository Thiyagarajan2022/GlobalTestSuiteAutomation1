﻿//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT EnvParams
//USEUNIT WorkspaceUtils
//UNEUNIT Datasheet


var unitName, testCase, execute, description;
var TestCase_ID, functions, Execute, Data;
var Opcolist = [];
var testOpco = "";
var globalTime,sheet;
var datasheetPath = [];
var workBook = "";
var folderName = null;
var testCaseId = null;
function executeTestCases(){
var instance = EnvParams.getEnvironment();
var businessFlow = EnvParams.getBusinessFlow();
globalTime = WorkspaceUtils.StartTime();
//Log.Message(EnvParams.instanceData);
//Log.Message(EnvParams.Country)
//Log.Message(EnvParams.testcase)
//Log.Message(EnvParams.TestingType)
//Log.Message(EnvParams.OpcoNum)
//Log.Message(EnvParams.Lang_Jenk)
//Log.Message(EnvParams.Opco)
//Log.Message(EnvParams.Language)
var ReportDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y_%H.%M.%S");
Log.Message("HTML REPORT PATH::"+Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate);
Log.Message("-----------------------------------------------------");
ReportUtils.createConsolidatedReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", "ConsolidatedReport");
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
var rowcount = ExcelUtils.getRowCount()-1;
var excelRow=1;
Opcolist = [];

//Checking whether need to execute ALL TestCase or NOT
if((EnvParams.testcase==null)||(EnvParams.testcase=="")||(EnvParams.testcase=="ALL")){
var excelName = EnvParams.getEnvironment();
workBook = Project.Path+excelName;
//Checking whether need to execute ALL TestCase for ALL OPCO'S or NOT
  if(EnvParams.OpcoNum=="ALL"){ 
Opcolist = columnCount(workBook,"Server Details");
  }else{     
   Opcolist [0] = EnvParams.Opco;
  }
 ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);

for(var OpID=0;OpID<Opcolist.length;OpID++){
  EnvParams.Opco = Opcolist[OpID];
  testOpco = Opcolist[OpID];
  excelRow=0
  Log.Message("TestRunner :"+EnvParams.Opco)
//Getting Language  

/*
var sheetName = "Server Details";
ExcelUtils.setExcelName(workBook, sheetName, true);  
Language = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
//Log.Message(Language)
if(EnvParams.Lang_Jenk==null){
if((Language=="")||(Language==null)){
//Language = EnvParams.Language;
}
else
EnvParams.Language = Language;
}
Log.Message(EnvParams.Language)
EnvParams.SetLanguage(Language)
*/



while(excelRow<=rowcount){
folderName = Opcolist[OpID];


ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
unitName = ExcelUtils.getAllRowDatas("UnitName",excelRow);

testCase = ExcelUtils.getAllRowDatas("TestCases",excelRow);

description = ExcelUtils.getAllRowDatas("Description",excelRow);

execute = ExcelUtils.getAllRowDatas("Execute",excelRow);

//Login for each Opco 
if(execute.toUpperCase()=="YES"){
reportName = "Report_"+EnvParams.Opco+"_"+unitName+"_"+testCase+".html"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest(unitName+" "+testCase, description)
var FolderID = Log.CreateFolder(Opcolist[OpID]+"_"+unitName);
Log.PushLogFolder(FolderID);
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"JIRA_Details",true)
testCaseId = ExcelUtils.getRowDatas(unitName,EnvParams.Country)

if(OpID==0){ 
reportName = "Report_"+EnvParams.Opco+"_ServerConfiguration.html"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest("ServerConfiguration login", "Login using given Credentials")
var FolderID = Log.CreateFolder("ServerConfiguration");
Log.PushLogFolder(FolderID);
Runner.CallMethod("ServerConfig.login");
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();  
}else{
reportName = "Report_"+EnvParams.Opco+"_Login_login.html"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest("Login login", "Login using given Credentials")
var FolderID = Log.CreateFolder("Login");
Log.PushLogFolder(FolderID);
Runner.CallMethod("Login.login");
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();
}



Runner.CallMethod(unitName+"."+testCase);
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();
Runner.CallMethod("JIRA.JIRAUpdate",folderName,testCaseId);
}
 
excelRow++;
} 

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
Delay(3000);
//  Log.Message("Maconomy is Already in Running")
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
}
}









else{ 
  
if(EnvParams.OpcoNum=="ALL"){ 
excelName = EnvParams.getEnvironment();
workBook = Project.Path+excelName;
Opcolist = columnCount(workBook,"Server Details");
  }else{ 
  excelName = EnvParams.getEnvironment();
  workBook = Project.Path+excelName;
  Opcolist [0] = EnvParams.Opco;
  }
  
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
for(var OpID=0;OpID<Opcolist.length;OpID++){
EnvParams.Opco = Opcolist[OpID];
testOpco = Opcolist[OpID];
Log.Message("Test Runner :"+EnvParams.Opco);
//Getting Language

/*
Log.Message(workBook)
var sheetName = "Server Details";
ExcelUtils.setExcelName(workBook, sheetName, true);  
Language = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
Log.Message(Language);
if(EnvParams.Lang_Jenk==null){
if((Language=="")||(Language==null)){
//Language = EnvParams.Language;
}
else
EnvParams.Language = Language;
}
EnvParams.SetLanguage(Language)
*/

////Login for each Opco
folderName = Opcolist[OpID];
if(OpID==0){ 
reportName = "Report_"+EnvParams.Opco+"_ServerConfiguration.html"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest("ServerConfiguration login", "Login using given Credentials")
var FolderID = Log.CreateFolder("ServerConfiguration");
Log.PushLogFolder(FolderID);
Runner.CallMethod("ServerConfig.login");
Log.PopLogFolder();  
ReportUtils.report.endTest(test);
ReportUtils.report.flush();
}else{
reportName = "Report_"+EnvParams.Opco+"_Login_login.html"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest("Login login", "Login using given Credentials")
var FolderID = Log.CreateFolder("Login");
Log.PushLogFolder(FolderID);
Runner.CallMethod("Login.login");
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();
}




ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
testCase = ExcelUtils.getRowDatas(EnvParams.testcase,"TestCases");
description = ExcelUtils.getRowDatas(EnvParams.testcase,"Description");
reportName = "Report_"+EnvParams.Opco+"_"+EnvParams.testcase+"_"+testCase+".html"
//Log.Message(reportName)
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
ReportUtils.createTest(EnvParams.testcase+" "+testCase, description)
var FolderID = Log.CreateFolder(Opcolist[OpID]+"_"+EnvParams.testcase);
Log.PushLogFolder(FolderID);
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"JIRA_Details",true)
testCaseId = ExcelUtils.getRowDatas(unitName,EnvParams.Country);
//Log.Message(EnvParams.testcase);
//Log.Message(testCase);
Runner.CallMethod(EnvParams.testcase+"."+testCase);
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();
Runner.CallMethod("JIRA.JIRAUpdate",folderName,testCaseId);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
//  Log.Message("Maconomy is Already in Running")
Delay(3000);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);

}
//}
}
ReportUtils.reportConsolidated.endTest(testConsolidated);
ReportUtils.reportConsolidated.flush();
}

function getExcelData(){ 
  excelData =[];  
  var colsList = [];
var workBook = Project.Path+EnvParams.path+"Book1.xlsx";
var sheetName = "Datasheets";

  var xlDriver = DDT.ExcelDriver(workBook, sheetName, true);
var id =0;
 for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
   colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
 }
  while (!DDT.CurrentDriver.EOF()) {
  datasheetPath[id] = Project.Path+EnvParams.path+xlDriver.Value(colsList[0]).toString().trim()+".xlsx";
  Log.Message( datasheetPath[id].substring(datasheetPath[id].lastIndexOf("\\")+1,datasheetPath[id].indexOf("."))); 
  id++;  
  xlDriver.Next();
  }    
  DDT.CloseDriver(xlDriver.Name);
  for(var i=0; i<datasheetPath.length;i++)
  { 
    Log.Message(datasheetPath[i]);
  }  
  
  return  datasheetPath; 
}

function CalculatingRootFilesNumber()
{

STIME = WorkspaceUtils.StartTime();
Log.Message(STIME)
   var foundFiles, aFile;
  foundFiles = aqFileSystem.FindFiles("C:\\Users\\674087\\Documents\\TestComplete 14 Projects\\WppRegression\\WppRegPack\\Testing Type\\Regression\\India", "*.xlsx");
  if (!strictEqual(foundFiles, null))
    while (foundFiles.HasNext())
    {
      aFile = foundFiles.Next();
      Log.Message(aFile.Name);
    }
  else
    Log.Message("No files were found.");

}


function columnCount(excelName,sheet){ 
  var colsList = [];
  var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
  var i=0;
   for(var idx=1;idx<DDT.CurrentDriver.ColumnCount;idx++){   
   colsList[i] = DDT.CurrentDriver.ColumnName(idx);
   Log.Message("Column :"+colsList[i]);
   i++;
 }
// Log.Message(colsList.length);
// Log.Message(colsList[0])
// Log.Message(colsList[1])
DDT.CloseDriver(xlDriver.Name);
 return colsList;
}