﻿//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT EnvParams
//USEUNIT WorkspaceUtils
//UNEUNIT Datasheet
 
 
var unitName, testCase, execute, description,sTime,eTime,testName,JkinsName,Okta="";
var TestCase_ID, functions, Execute, Data;
var Opcolist = [];
var CountryList = [];
var testOpco = "";
var globalTime,sheet;
var datasheetPath = [];
var workBook = "";
var folderName = null;
var testCaseId = null;
var releasename =null;
var cyclename =null;
var workDir = "";
var packedResults = "";
var reportName = "";
var archivePath = "";
var exeResults ="";
var automationStat_file = "";
var testCase_Stat_updated_flag;
var TempUnit = "";
var JiraUpdate = true;
var JiraStat = true;
//var JfolderName = "";

function executeTestCases(){
var instance = EnvParams.getEnvironment();
var businessFlow = "";

//var businessFlow = EnvParams.getBusinessFlow();
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
var automationStat_Date = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y");
Log.Message("HTML REPORT PATH::"+Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate);
Log.Message("-----------------------------------------------------");
ReportUtils.createConsolidatedReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\", "ConsolidatedReport");
//ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
//var rowcount = ExcelUtils.getRowCount()-1;
//var excelRow=1;
var rowcount = "";
var excelRow="";
Opcolist = [];
CountryList = [];

if((EnvParams.CountryList=="ALL")){ //Checking whether need to execute ALL TestCase for ALL Country or NOT
   CountryList = getRowDatas(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"CountryMapping",EnvParams.instanceData);
  }else{  
  if((EnvParams.CountryList!=null)&&(EnvParams.CountryList!="")){
  if(EnvParams.CountryList.indexOf(",")!=-1){
   CountryList = EnvParams.CountryList.split(",");
   }
   else{
   CountryList [0] = EnvParams.CountryList;
   } 
   }
  else{
   CountryList [0] = EnvParams.Country;
   }       
  }
  
for(var contyID =0;contyID<CountryList.length;contyID++){
  EnvParams.Country = CountryList[contyID];
//  Log.Message("CountryList[contyID] :"+CountryList[contyID]);
  setPath(CountryList[contyID]);
//Log.Message("Path :"+EnvParams.path);
businessFlow = EnvParams.getBusinessFlow();
//ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
if(TestingType.toUpperCase()=="SMOKE")
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"Smoke");
else
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"GlobalTestCase");
var rowcount = ExcelUtils.getRowCount()-1;
var excelRow=1;
if((EnvParams.testcase==null)||(EnvParams.testcase=="")||(EnvParams.testcase=="ALL")||(EnvParams.TestingType=="Full_Regression")||(EnvParams.TestingType=="Smoke")){   //Checking whether need to execute ALL TestCase or NOT
var excelName = EnvParams.path;
workBook = Project.Path+excelName;
  
if(EnvParams.OpcoNum=="ALL"){ //Checking whether need to execute ALL TestCase for ALL Country or NOT
   Opcolist = columnCount(workBook,"Server Details");
  }else{ 
 if((EnvParams.OpcoNum!=null)&&(EnvParams.OpcoNum!="")){
 if(EnvParams.OpcoNum.indexOf(",")!=-1){
   Opcolist = EnvParams.OpcoNum.split(",");
   }
 else{
  Opcolist [0] = EnvParams.OpcoNum;    
  } 
   }
  else{
  Opcolist [0] = EnvParams.Opco;    
  }  
  }
  
// ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
 if(TestingType.toUpperCase()=="SMOKE")
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"Smoke");
else
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"GlobalTestCase");
var server = true;
var nxtID = -1;
for(var OpID=0;OpID<Opcolist.length;OpID++){
  EnvParams.Opco = Opcolist[OpID];
  testOpco = Opcolist[OpID];
  excelRow=0
  Log.Message("TestRunner :"+EnvParams.Opco)
  
  var Coun_Opco = getRowOPco(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"OpcoMapping",EnvParams.Country.toUpperCase(),testOpco);
if(!Coun_Opco){ 
//  Log.Warning("Opco Number :"+testOpco+" is not in Country :"+EnvParams.Country);
  continue;
}
//Getting Language  


//var app = Sys.OleObject("Excel.Application");
//var book = app.Workbooks.Add();
//app.Visible = "True";

//var columnCount = book.ActiveSheet.UsedRange.Columns.Count;
//var rowCount = book.ActiveSheet.UsedRange.Rows.Count;



var executionTime;

while(excelRow<=rowcount){
folderName = Opcolist[OpID];


//ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
if(TestingType.toUpperCase()=="SMOKE")
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"Smoke");
else
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"GlobalTestCase");

JkinsName = ExcelUtils.getAllRowDatas("Jenkins Name",excelRow);
unitName = ExcelUtils.getAllRowDatas("UnitName",excelRow);
testCase = ExcelUtils.getAllRowDatas("TestCases",excelRow);
moduleName = ExcelUtils.getAllRowDatas("ModuleName",excelRow);
description = moduleName;

if(TestingType.toUpperCase()=="SMOKE")
execute = ExcelUtils.getAllRowDatas("Execute",excelRow);
else
execute = ExcelUtils.getAllRowDatas(businessFlow,excelRow);

if(execute.toUpperCase()=="YES"){   //Login for each Opco 

ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"JIRA_Details",true)
testCaseId = ExcelUtils.getRowDatas(JkinsName,EnvParams.Country)
releasename  = ExcelUtils.getRowDatas("Current Release Name",EnvParams.Country)
//cyclename  = ExcelUtils.getRowDatas("Current Cycle Name",EnvParams.Country)
ExcelUtils.setExcelName(workBook, "Server Details", true);
cyclename = ExcelUtils.getRowDatas("JIRA Cycle Name",EnvParams.Opco)
folderName = ExcelUtils.getRowDatas("JIRA Folder Name",EnvParams.Opco)
Okta = ExcelUtils.getRowDatas("Okta",EnvParams.Opco);

if(server){ 
JiraUpdate = true;
JiraStat = true;
//      reportName = "Report_"+EnvParams.Opco+"_Login";
//      ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
//      var LworkDir = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\"+reportName+"\\";
//      var LpackedResults = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\";
//    //ReportUtils.createTest("Login login", "Login using given Credentials")
//      ReportUtils.createTest("Login", "Login using given Credentials")
//
//      var FolderID = Log.CreateFolder("Login");
//      Log.PushLogFolder(FolderID);
//     Runner.CallMethod("Login.login");
//      Log.PopLogFolder();
//      ReportUtils.report.endTest(test);
//      ReportUtils.report.flush();
//      fileList = slPacker.GetFileListFromFolder(LworkDir);
//      archivePath = LpackedResults +reportName;
//// Packes the resutls
//if (slPacker.Pack(fileList, LworkDir, archivePath))
//      Log.Message("Files compressed successfully."); 
   
}

server = false;

testName = unitName;

testCase_Stat_updated_flag = false;

reportName = "Report_"+EnvParams.Opco+"_"+unitName;
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\", reportName);

exeResults = Project.Path+TextUtils.GetProjectValue("ReportPath");
automationStat_file = exeResults+"RunTime_statistics_"+automationStat_Date+".xlsx";
workDir = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\"+reportName+"\\";
packedResults = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\";

ReportUtils.createTest(unitName, description);
JiraUpdate = true;
JiraStat = true;
// capture StartTime
sTime = new Date();
TextUtils.writeLog(unitName +" Execution Started Time :"+sTime); 


var FolderID = Log.CreateFolder(Opcolist[OpID]+"_"+unitName);
Log.PushLogFolder(FolderID);
Log.Message(unitName+"."+testCase);
Runner.CallMethod(unitName+"."+testCase);
Log.PopLogFolder();
TextUtils.writeLog(unitName+" Completed Successfully");

ReportUtils.report.endTest(test);
ReportUtils.report.flush();

fileList = slPacker.GetFileListFromFolder(workDir);
archivePath = packedResults + reportName;
Delay(5000);
if (slPacker.Pack(fileList, workDir, archivePath))
  Log.Message("Files compressed successfully");
  
Runner.CallMethod("JIRA.JIRAUpdate");

// capture EndTime
eTime = new Date();;
TextUtils.writeLog(unitName +" Execution Ended Time :"+eTime); 

// Verify Statistics file exists or not. If not create it.
if(!aqFile.Exists(automationStat_file))
 ExcelUtils.create_AutomationStat_Excel(automationStat_file);  

// Calculate RunTime and publish in Excel 

executionTime = 0;    
executionTime = WorkspaceUtils.timeDifference(sTime, eTime)   
ExcelUtils.writeTo_AutomationStat_Excel(automationStat_file,moduleName,JkinsName,EnvParams.Opco,executionTime);
testCase_Stat_updated_flag=true;
}
 
excelRow++;

} 


var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
Delay(3000);

 }
}
else{ 
  
businessFlow = EnvParams.getBusinessFlow(); 
if((EnvParams.OpcoNum==null)||(EnvParams.OpcoNum=="")||(EnvParams.OpcoNum=="ALL")){ //Checking whether need to execute ALL TestCase for ALL Country or NOT
excelName = EnvParams.path;
workBook = Project.Path+excelName;   
Opcolist = columnCount(workBook,"Server Details");
  }else{ 
    
  if(EnvParams.OpcoNum.indexOf(",")!=-1){
  excelName = EnvParams.path;
  workBook = Project.Path+excelName;
  Opcolist = EnvParams.OpcoNum.split(",");
   }
   else{ 
   excelName = EnvParams.path;
   workBook = Project.Path+excelName;
   Opcolist [0] = EnvParams.OpcoNum;
   }
  } 
  
var testList = [];
  if(EnvParams.testcase.indexOf(",")!=-1){
  testList = EnvParams.testcase.split(",");
   }
   else{ 
   testList [0] = EnvParams.testcase;
   }
  
if(TestingType.toUpperCase()=="SMOKE")
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"Smoke");
else
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"GlobalTestCase");
for(var OpID=0;OpID<Opcolist.length;OpID++){
EnvParams.Opco = Opcolist[OpID];
testOpco = Opcolist[OpID];
Log.Message("Test Runner :"+EnvParams.Opco);

var Coun_Opco = getRowOPco(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"OpcoMapping",EnvParams.Country.toUpperCase(),testOpco);
if(!Coun_Opco){ 
//  Log.Warning("Opco Number :"+testOpco+" is not in Country :"+EnvParams.Country);
  continue;
}


ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"JIRA_Details",true)
testCaseId = ExcelUtils.getRowDatas(testList[0],EnvParams.Country);
Log.Message("testList[0]"+testCaseId)
releasename  = ExcelUtils.getRowDatas("Current Release Name",EnvParams.Country)
ExcelUtils.setExcelName(workBook, "Server Details", true);
cyclename = ExcelUtils.getRowDatas("JIRA Cycle Name",EnvParams.Opco)
folderName = ExcelUtils.getRowDatas("JIRA Folder Name",EnvParams.Opco)
JiraUpdate = true;
JiraStat = true;
//folderName = Opcolist[OpID];   //Login for each Opco


if(OpID==0){ 
reportName = "Report_"+EnvParams.Opco+"_Login";
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
var LworkDir = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\"+reportName+"\\";
var LpackedResults = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\";
ReportUtils.createTest("Login", "Login using given Credentials")
var FolderID = Log.CreateFolder("Login");
Log.PushLogFolder(FolderID);
Runner.CallMethod("Login.login");
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();  
Delay(4000);
fileList = slPacker.GetFileListFromFolder(LworkDir);
archivePath = LpackedResults + reportName;
// Packes the resutls
if (slPacker.Pack(fileList, LworkDir, archivePath))
  Log.Message("Files compressed successfully.");

}else{
reportName = "Report_"+EnvParams.Opco+"_Login"
ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
//ReportUtils.createTest("Login login", "Login using given Credentials")
var LworkDir = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\"+reportName+"\\";
var LpackedResults = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\";
ReportUtils.createTest("Login", "Login using given Credentials")
var FolderID = Log.CreateFolder("Login");
Log.PushLogFolder(FolderID);
Runner.CallMethod("Login.login");
Log.PopLogFolder();
ReportUtils.report.endTest(test);
ReportUtils.report.flush();
fileList = slPacker.GetFileListFromFolder(LworkDir);
archivePath = LpackedResults + reportName;
// Packes the resutls
if (slPacker.Pack(fileList, LworkDir, archivePath))
  Log.Message("Files compressed successfully.");
}



//var app = Sys.OleObject("Excel.Application");
//app.Visible = "True";
//var book = app.Workbooks.Open("Test");
//var sheet = book.Sheets.Item("First");
//var columnCount = sheet.UsedRange.Columns.Count;
//var rowCount = sheet.UsedRange.Rows.Count;
//var col =0;
//var row = 0;

testCase_Stat_updated_flag = false;

for(var tL=0;tL<testList.length;tL++){
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"JIRA_Details",true)
testCaseId = ExcelUtils.getRowDatas(testList[tL],EnvParams.Country);
releasename  = ExcelUtils.getRowDatas("Current Release Name",EnvParams.Country)
ExcelUtils.setExcelName(workBook, "Server Details", true);
Log.Message("workbook Name"+workBook)
cyclename = ExcelUtils.getRowDatas("JIRA Cycle Name",EnvParams.Opco)
folderName = ExcelUtils.getRowDatas("JIRA Folder Name",EnvParams.Opco)
Log.Message(cyclename)
Log.Message(folderName)
JiraUpdate = true;
JiraStat = true;

//ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),businessFlow);
if(TestingType.toUpperCase()=="SMOKE")
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"Smoke");
else
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("RunManagerPath"),"GlobalTestCase");
Log.Message("TestCase Name"+testList[tL])
JkinsName = testList[tL];
unitName = ExcelUtils.getRowDatas(testList[tL],"UnitName");
testCase = ExcelUtils.getRowDatas(testList[tL],"TestCases");
description = ExcelUtils.getRowDatas(testList[tL],"Module");
moduleName = description;
reportName = "Report_"+EnvParams.Opco+"_"+testList[tL];
Log.Message(testCase)
Log.Message(description)

exeResults = Project.Path+TextUtils.GetProjectValue("ReportPath");
automationStat_file = exeResults+"RunTime_statistics_"+automationStat_Date+".xlsx";

ReportUtils.createReport(Project.Path+TextUtils.GetProjectValue("ReportPath")+"\\"+"Report_"+ReportDate+"\\", reportName);
workDir = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\"+reportName+"\\";
packedResults = Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\";
//unitName = testList[tL];

ReportUtils.createTest(unitName, description);

// capture StartTime
sTime = new Date();
TextUtils.writeLog(unitName +" Execution Started Time :"+sTime); 


//sheet.Cells.Item(rowCount+1,  1).Value = testList(tL);
//sheet.Cells.Item(rowCount+1,  col).Value = sTime;

var FolderID = Log.CreateFolder(Opcolist[OpID]+"_"+unitName);
Log.PushLogFolder(FolderID);
Log.Message(unitName);
Log.Message(testCase);
Runner.CallMethod(unitName+"."+testCase);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog(unitName+" PASSED and Completed Successfully");

eTime = WorkspaceUtils.StartTime();
TextUtils.writeLog(unitName +" Execution End Time :"+eTime); 

//sheet.Cells.Item(rowCount+1,  1).Value = testList(tL);
//sheet.Cells.Item(rowCount+1,  col).Value = eTime;

ReportUtils.report.endTest(test);
ReportUtils.report.flush();

//book.Save();
//app.Quit();

fileList = slPacker.GetFileListFromFolder(workDir);
archivePath = packedResults + reportName;
aqUtils.Delay(4000, "Updating Result in JIRA");
// Packes the resutls
//if (slPacker.Pack(fileList, workDir, archivePath))
//  Log.Message("Files compressed successfully.");

//Runner.CallMethod("JIRA.JIRAUpdate",folderName,testCaseId,releasename,cyclename);
Runner.CallMethod("JIRA.JIRAUpdate");

// capture EndTime
eTime = new Date();;
TextUtils.writeLog(unitName +" Execution Ended Time :"+eTime); 

// Verify Statistics file exists or not. If not create it.
if(!aqFile.Exists(automationStat_file))
 ExcelUtils.create_AutomationStat_Excel(automationStat_file);  

// Calculate RunTime and publish in Excel 

executionTime = 0;    
executionTime = WorkspaceUtils.timeDifference(sTime, eTime)   
ExcelUtils.writeTo_AutomationStat_Excel(automationStat_file,moduleName,testList[tL],EnvParams.Opco,executionTime);
testCase_Stat_updated_flag=true;


}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();

aqUtils.Delay(3000, "Closing Maconomy");
    Sys.Desktop.KeyDown(0x12); //Alt  //  Log.Message("Maconomy is Already in Running")
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);

}
}
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
//   Log.Message("Column :"+colsList[i]);
   i++;
 }
// Log.Message(colsList.length);
// Log.Message(colsList[0])
// Log.Message(colsList[1])
DDT.CloseDriver(xlDriver.Name);
 return colsList;
}

function getRowDatas(excelName,sheet,column)
{

//Log.Message("excelName :"+excelName);
//Log.Message("sheet :"+sheet);
var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var rowList = [];
 var temp ="";

     while (!DDT.CurrentDriver.EOF()) {
//    Log.Message("Colunm :"+xlDriver.Value(0).toString().trim())

        try{
          
         rowList[id] = xlDriver.Value(column).toString().trim();
         id++;
         }
        catch(e){
        temp = "";
        }

//      Log.Message("temp :"+temp);


    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
//      Log.Message("rowList :"+rowList);
     return rowList;
}


function getRowOPco(excelName,sheet,column,OpID)
{

//Log.Message("excelName :"+excelName);
//Log.Message("sheet :"+sheet);
//Log.Message("column :"+column);
//Log.Message("OpID :"+OpID);
var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var rowList = [];
 var temp =false;

     while (!DDT.CurrentDriver.EOF()) {
//    Log.Message("Colunm :"+xlDriver.Value(0).toString().trim())

        try{
          if(OpID==xlDriver.Value(column).toString().trim()){
          temp = true;
          break;
          }
         }
        catch(e){
//        temp = "";
        }

//      Log.Message("temp :"+temp);

//      Log.Message("temp :"+temp);
    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
//      Log.Message("temp :"+temp);     
     return temp;
}
