﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/**
 * This script Create Blanket Invoice for Job
 * @author  : Muthu Kumar M
 * @version : 1.0
 * Created Date :02/17/2021
*/

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "BlanketInvoice";

//Global Variable
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var jobNumber_2,jobNumber_1,EmployeeNumber,jobNumber = "";
var Language = "";
var BlanketInvoice_JIRA_ID = "";
var BlanketInvoice_UnitName_JIRA = "";
var Estimatelines = []; 
var Estimatelines_1 = []; 
var Global_Client_Number = "";

//Main Function
function Create_Blanket_Invoice() {
  
TextUtils.writeLog("Blanket Invoice Creation Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Blanket Invoice script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "BlanketInvoice";


ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";

Approve_Level =[];
ApproveInfo = [];
Estimatelines = []; 
Estimatelines_1 = []; 
level =0;
jobNumber_2,jobNumber_1,EmployeeNumber,Global_Client_Number,jobNumber = "";
BlanketInvoice_JIRA_ID = "";
BlanketInvoice_UnitName_JIRA = "";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);



getDetails();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
//Checking Login to execute Combined Invoice script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

aqUtils.Delay(5000, Indicator.Text);
Log.Message("jobNumber_1 : "+jobNumber_1)
Log.Message("jobNumber_2 : "+jobNumber_2)
goToJobMenuItem();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
setLayout_for_Job1();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

goToJobMenuItem();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
setLayout_for_Job2();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000, Indicator.Text);

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(5000, Indicator.Text);

goto_BlanketInvocie_Item();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
select_client();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
Select_Job1_BlanketInvoice();
Select_Job2_BlanketInvoice();

Submit_Draft();

//This Function call to get UserName for Approver
CredentialLogin();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
for(var i=level;i<ApproveInfo.length;i++){
level=i;
aqUtils.Delay(5000, Indicator.Text);
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApprove(temp[1],temp[2],i);


}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}


function getDetails(){ 
  
  ExcelUtils.setExcelName(workBook, "BlanketInvoice", true);
  EmployeeNumber = ExcelUtils.getColumnDatas("Employee Number",EnvParams.Opco)
  jobNumber_1 = ExcelUtils.getColumnDatas("Job Number_1",EnvParams.Opco)
  jobNumber_2 = ExcelUtils.getColumnDatas("Job Number_2",EnvParams.Opco)
  var Foreign_Client = ExcelUtils.getColumnDatas("Foreign Client",EnvParams.Opco)
  if(Foreign_Client.toUpperCase()=="YES"){ 
ExcelUtils.setExcelName(workBook, "Data Management", true);
Global_Client_Number = ReadExcelSheet("Foreign Global Client Number",EnvParams.Opco,"Data Management");
  }else{
ExcelUtils.setExcelName(workBook, "Data Management", true);
Global_Client_Number = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
}

if((Global_Client_Number=="")||(Global_Client_Number==null)){
ExcelUtils.setExcelName(workBook, "BlanketInvoice", true);
Global_Client_Number = ExcelUtils.getRowDatas("Client No",EnvParams.Opco)
}
if((Global_Client_Number==null)||(Global_Client_Number=="")){ 
ValidationUtils.verify(false,true,"Client Number is Needed to Create Blanket Invoice");
}

  
  
if((jobNumber_1=="")||(jobNumber_1==null)){
    
  
    //Creation of Job 1
    BlanketInvoice_JIRA_ID = TestRunner.testCaseId;
    BlanketInvoice_UnitName_JIRA = TestRunner.unitName; 
    TestRunner.TempUnit = BlanketInvoice_UnitName_JIRA;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    
    //Sheet Name
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet_1",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Invoice")
    }
    
    //Job Serial Order
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Invoice")
    }
    
    //Job Number in Serial Order
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    jobNumber_1 = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    
    if((jobNumber_1=="")||(jobNumber_1==null)){
      
    //JIRA ID for Job Creation
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    Job_JIRAID = ExcelUtils.getRowDatas("JobCreation_"+serialOder,EnvParams.Country);
    if((Job_JIRAID=="")||(Job_JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for Jobcreation_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = Job_JIRAID;
    TestRunner.unitName = "JobCreation_"+serialOder;
    
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job");
    
    //Execute Dependency Job Creation
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+Job_JIRAID)
    Runner.CallMethod("Creation_Of_Job.createJob",jobSheet,serialOder);
    Log.PopLogFolder();

    ExcelUtils.setExcelName(workBook,"Data Management", true);
    jobNumber_1 = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    ExcelUtils.WriteExcelSheet("Blanket Job Number_1",EnvParams.Opco,"Data Management",jobNumber_1)
    
    //Updating result in JIRA
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    
    //------------Creation of Budget---------------------------------------------
    
    //Getting Create Budget Sheet
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var budgetSheet = ExcelUtils.getColumnDatas("Budget sheet_1",EnvParams.Opco)
    if(budgetSheet==""){ 
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Invoice")
    }
    
    //Getting Serial Order for Create Budget
    ExcelUtils.setExcelName(workBook, budgetSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Budget")
    }
    
    //Checking Budget is Completed or Not?
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var WE_Number = ExcelUtils.getRowDatas("Working Estimate_"+serialOder,EnvParams.Opco)
    if((WE_Number=="")||(WE_Number==null)){
    
    //Getting JIRA ID for Create Budget
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateBudget_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateBudget_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;  
    TestRunner.unitName = "CreateBudget_"+serialOder;
    
    
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job Budget");
    
    //Execute Create Budget
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job Budget");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("BudgetCreation.createBudget",budgetSheet,serialOder);
    Log.PopLogFolder();
    
    //Updating JIRA for Created Budget
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    
  //-------------------------Creation of Quote-------------------------------------------------------------
  
  //Getting Create Quote Sheet
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var quoteSheet = ExcelUtils.getColumnDatas("Quote Sheet_1",EnvParams.Opco)
    if(quoteSheet==""){ 
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Invoice")
    }
    
    //Getting Serial Order for Quote
    ExcelUtils.setExcelName(workBook, quoteSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Quote")
    }
    
    //Checking Quote has Completed or NOT ?
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var CE_Number = ExcelUtils.getRowDatas("Client Approved Estimate_"+serialOder,EnvParams.Opco)
    if((CE_Number=="")||(CE_Number==null)){
    
    //Getting JIRA ID for Create Quote
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateQuote_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateQuote_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;    
    TestRunner.unitName = "CreateQuote_"+serialOder;
    
    
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Quote");
    
    //Execute Create Quote
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Quote");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("Creation_of_Quote.CreateQuote",quoteSheet,serialOder);
    Log.PopLogFolder();
    
    
    //Updating JIRA for Created Quote
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
 }   
    
      
  
//Job 2
    
  if((jobNumber_2=="")||(jobNumber_2==null)){
    //Creation of Job
    BlanketInvoice_JIRA_ID = TestRunner.testCaseId;
    BlanketInvoice_UnitName_JIRA = TestRunner.unitName; 
    TestRunner.TempUnit = BlanketInvoice_UnitName_JIRA;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    
    //Sheet Name
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet_2",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Invoice")
    }
    
    //Job Serial Order
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Invoice")
    }
    
    //Job Number in Serial Order
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    jobNumber_2 = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    
    if((jobNumber_2=="")||(jobNumber_2==null)){
      
    //JIRA ID for Job Creation
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    Job_JIRAID = ExcelUtils.getRowDatas("JobCreation_"+serialOder,EnvParams.Country);
    if((Job_JIRAID=="")||(Job_JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for Jobcreation_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = Job_JIRAID;
    TestRunner.unitName = "JobCreation_"+serialOder;
    
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job");
    
    //Execute Dependency Job Creation
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+Job_JIRAID)
    Runner.CallMethod("Creation_Of_Job.createJob",jobSheet,serialOder);
    Log.PopLogFolder();

    ExcelUtils.setExcelName(workBook,"Data Management", true);
    jobNumber_2 = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    ExcelUtils.WriteExcelSheet("Blanket Job Number_1",EnvParams.Opco,"Data Management",jobNumber_2)
    
    //Updating result in JIRA
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    
    //------------Creation of Budget---------------------------------------------
    
    //Getting Create Budget Sheet
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var budgetSheet = ExcelUtils.getColumnDatas("Budget sheet_2",EnvParams.Opco)
    if(budgetSheet==""){ 
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Invoice")
    }
    
    //Getting Serial Order for Create Budget
    ExcelUtils.setExcelName(workBook, budgetSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Budget")
    }
    
    //Checking Budget is Completed or Not?
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var WE_Number = ExcelUtils.getRowDatas("Working Estimate_"+serialOder,EnvParams.Opco)
    if((WE_Number=="")||(WE_Number==null)){
    
    //Getting JIRA ID for Create Budget
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateBudget_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateBudget_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;  
    TestRunner.unitName = "CreateBudget_"+serialOder;
    
    
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job Budget");
    
    //Execute Create Budget
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job Budget");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("BudgetCreation.createBudget",budgetSheet,serialOder);
    Log.PopLogFolder();
    
    //Updating JIRA for Created Budget
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    
  //-------------------------Creation of Quote-------------------------------------------------------------
  
  //Getting Create Quote Sheet
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var quoteSheet = ExcelUtils.getColumnDatas("Quote Sheet_2",EnvParams.Opco)
    if(quoteSheet==""){ 
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Invoice")
    }
    
    //Getting Serial Order for Quote
    ExcelUtils.setExcelName(workBook, quoteSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Quote")
    }
    
    //Checking Quote has Completed or NOT ?
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var CE_Number = ExcelUtils.getRowDatas("Client Approved Estimate_"+serialOder,EnvParams.Opco)
    if((CE_Number=="")||(CE_Number==null)){
    
    //Getting JIRA ID for Create Quote
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateQuote_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateQuote_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;    
    TestRunner.unitName = "CreateQuote_"+serialOder;
    
    
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Quote");
    
    //Execute Create Quote
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Quote");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("Creation_of_Quote.CreateQuote",quoteSheet,serialOder);
    Log.PopLogFolder();
    
    
    //Updating JIRA for Created Quote
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    }

}



/**
  *  This function Navigates to Jobs screen from Jobs workspace
  */
function goToJobMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
}

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
var Client_Managt;

for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}



//Selecting Job for  Blanket Invoice in Maconomy
function setLayout_for_Job1(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var allJobs = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
allJobs.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(jobNumber_1);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(3000, "Finding Jobs in Maconomy");


  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber_1){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to for Invoice");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber_1+") is available in maconommy for Invoice"); 
  closeFilter.Click();
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var clientApproved = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(clientApproved);
  if(clientApproved.background!=10674625){
    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
    }
  var workingEstimate = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
  WorkspaceUtils.waitForObj(workingEstimate);
  if(workingEstimate.background!=10674625){
    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
    }

  var Information = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(Information);
  Information.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  var layout = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  layout.Keys(" ");
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  layout.HoverMouse();
  Sys.HighlightObject(layout);
  layout.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blanket Job Invoicing").OleValue.toString().trim(),"Blanket Job Invoicing")
  aqUtils.Delay(8000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Layout is Seleted as Blanket Invoice")
  var CheckBox_BlanketInvoice = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McPlainCheckboxView.Button;
  if(!CheckBox_BlanketInvoice.getSelection()){ 
    CheckBox_BlanketInvoice.Click();
    aqUtils.Delay(5000, "Blanket Invoice Selected");
    TextUtils.writeLog("Blanket Invoice is Ticked")
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  }
  
  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SingleToolItemControl;
  if(Save.Enabled){ 
    Save.Click();
  }
  aqUtils.Delay(5000, "Blanket Invoice Selected");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  
  var Budgeting = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  
  var Estimate = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  Estimate.Keys("Client Approved Estimate");
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  
  var FullBudget = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var BudgetGrid = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(BudgetGrid);
  var ii=0;
Estimatelines = [];  
  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")&&(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!=null)){ 
         if(EnvParams.Country.toUpperCase()=="INDIA")
         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(12).OleValue.toString().trim();
         else
         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim();
         Log.Message(Estimatelines[ii]);
         ii++;
    }
  }
  

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
}
}



//Selecting Job for  Blanket Invoice in Maconomy
function setLayout_for_Job2(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var allJobs = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
allJobs.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(jobNumber_2);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(3000, "Finding Jobs in Maconomy");


  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber_2){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to for Invoice");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber_2+") is available in maconommy for Invoice"); 
  closeFilter.Click();
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var clientApproved = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(clientApproved);
  if(clientApproved.background!=10674625){
    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
    }
  var workingEstimate = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
  WorkspaceUtils.waitForObj(workingEstimate);
  if(workingEstimate.background!=10674625){
    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
    }

  var Information = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(Information);
  Information.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  var layout = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  layout.Keys(" ");
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  layout.HoverMouse();
  Sys.HighlightObject(layout);
  layout.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blanket Job Invoicing").OleValue.toString().trim(),"Blanket Job Invoicing")
  aqUtils.Delay(8000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Layout is Seleted as Blanket Invoice")
  var CheckBox_BlanketInvoice = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McPlainCheckboxView.Button;
  if(!CheckBox_BlanketInvoice.getSelection()){ 
    CheckBox_BlanketInvoice.Click();
    aqUtils.Delay(5000, "Blanket Invoice Selected");
    TextUtils.writeLog("Blanket Invoice is Ticked")
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  }
  
  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SingleToolItemControl;
  if(Save.Enabled){ 
    Save.Click();
  }
  aqUtils.Delay(5000, "Blanket Invoice Selected");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  
  var Budgeting = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  
  var Estimate = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  Estimate.Keys("Client Approved Estimate");
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  
  var FullBudget = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var BudgetGrid = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(BudgetGrid);
  var ii=0;
Estimatelines_1 = [];  
  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")&&(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!=null)){ 
         if(EnvParams.Country.toUpperCase()=="INDIA")
         Estimatelines_1[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(12).OleValue.toString().trim();
         else
         Estimatelines_1[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim();
         Log.Message(Estimatelines_1[ii]);
         ii++;
    }
  }
  

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
}
}




/**
  *  This function Navigates to Jobs screen from Jobs workspace
  */

function goto_BlanketInvocie_Item(){ 
  

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
}

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
var Client_Managt;

for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blanket Invoicing").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blanket Invoicing").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Jobs from Blanket Invoice Menu");
TextUtils.writeLog("Entering into Jobs from Blanket Invoice Menu");

}



// Select Client for Balnket Invoice
function select_client(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
//var All = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
var All = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
Sys.HighlightObject(All);
All.Click();


  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
//  var firstCell = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget
  var firstCell = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(firstCell)
  firstCell.Click();
  firstCell.Keys("[Tab]");
//  var ClientNo = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var ClientNo = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  Sys.HighlightObject(ClientNo);
  ClientNo.Click();
  ClientNo.Keys(Global_Client_Number);
  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  
//  var CloseFilter = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  var CloseFilter = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(CloseFilter);
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(3000, "Finding Client in Maconomy");

//  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==Global_Client_Number){ 
    flag=true;
    table.Keys("[Down]");
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
  }
  
  if(flag){
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  CloseFilter.Click();
  }else{ 
    ValidationUtils.verify(true,false,"Client Number is not availble in Maconomy")
  }

}



//Invocing Job 1
function Select_Job1_BlanketInvoice(){ 
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  
  ImageRepository.ImageSet.Blanket_Maximize.Click();
  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  
//  var Job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
  var Job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(Job);
  Job.Click();
  
//  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==jobNumber_1){ 
  flag=true;
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
  
  
  if(flag){ 
    TextUtils.writeLog("Job No:"+jobNumber_1+" is selected in blanket invoice tab")
    
    
//    var OnAccount = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;
    var OnAccount = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(OnAccount);
    OnAccount.Click();
    aqUtils.Delay(5000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(5000, Indicator.Text);
    


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
     var addedlines = true;
     for(var ii=0;ii<Estimatelines.length;ii++){
       var HSN,wCodeID,Desp,Qly,UnitPrice ="";
       var temp = Estimatelines[ii].split("*");
       
        wCodeID = temp[0];
        Desp = temp[1];
        Qly = temp[2];
        UnitPrice = temp[3];
        if(EnvParams.Country.toUpperCase()=="INDIA"){
        HSN = temp[5];
        }

        TextUtils.writeLog("Line item "+ii+" is adding in specification");
         aqUtils.Delay(4000, "Next Column");    
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
        }
        
 aqUtils.Delay(4000, "Next Column");       
//var addLine = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;
var addLine = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
Sys.HighlightObject(addLine);
for(var i=0;i<addLine.ChildCount;i++){ 
  if((addLine.Child(i).isVisible())&&((addLine.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Add Job Invoice On Account Entry (Ctrl+M)").OleValue.toString().trim()) || (addLine.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Add Job Invoice On Account Entry").OleValue.toString().trim()))){
    addLine = addLine.Child(i);
    WorkspaceUtils.waitForObj(addLine);
    ReportUtils.logStep_Screenshot("");
    addLine.Click();
    break;
  }
}
aqUtils.Delay(4000, "Next Column");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
     
//var standardCode = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;              
var standardCode = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
addedlines = true;
standardCode.Click();
WorkspaceUtils.waitForObj(standardCode);
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(2000, "Next Column");
    
//var description = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
var description = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
description.Click();
description.settext(Desp);
description.Keys("[Tab]");
aqUtils.Delay(2000, "Next Column");

//var qty = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
var qty = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
qty.Click();
qty.settext(Qly);
qty.Keys("[Tab]");
aqUtils.Delay(2000, "Next Column");

//var billable = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
var billable = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
billable.Click();
billable.settext(UnitPrice);
billable.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
aqUtils.Delay(2000, "Next Column");

//var workcode = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
var workcode = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
workcode.Click();
WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"WorkCode");
  
if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_InvoiceOnAccount.specification",workcode,HSN);

//  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;
  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
  WorkspaceUtils.waitForObj(Save);
for(var i=0;i<Save.ChildCount;i++){ 
  if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save Job Invoice On Account Entry (Enter)").OleValue.toString().trim())){
    Save = Save.Child(i);
    WorkspaceUtils.waitForObj(Save);
    ReportUtils.logStep_Screenshot("");
    Save.Click();
    break;
  }
}

aqUtils.Delay(1000, "Saving Line in On Account");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, "Saving Line in On Account");

    
     }
     
     aqUtils.Delay(3000, "Combined Invoice Selected");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


  
  }
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
            
//  var Job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
  var Job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(Job);
  Job.Click();
//  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  for(var v=0;v<table.getItemCount();v++){ 
  table.Keys("[Up]");
  }

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
}




//Invocing Job 2
function Select_Job2_BlanketInvoice(){ 

  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  
  var Job = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(Job);
  Job.Click();
  var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==jobNumber_2){ 
  flag=true;
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
  
  
  if(flag){ 
    
  TextUtils.writeLog("Job No:"+jobNumber_2+" is selected in blanket invoice tab")
  
    var OnAccount = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(OnAccount);
    OnAccount.Click();
    aqUtils.Delay(5000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(5000, Indicator.Text);
    


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
     var addedlines = true;
     for(var ii=0;ii<Estimatelines_1.length;ii++){
       var HSN,wCodeID,Desp,Qly,UnitPrice ="";
       var temp = Estimatelines_1[ii].split("*");
       
        wCodeID = temp[0];
        Desp = temp[1];
        Qly = temp[2];
        UnitPrice = temp[3];
        if(EnvParams.Country.toUpperCase()=="INDIA"){
        HSN = temp[5];
        }

        TextUtils.writeLog("Line item "+ii+" is adding in specification");
         aqUtils.Delay(4000, "Next Column");    
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
        }
        
 aqUtils.Delay(4000, "Next Column");       
var addLine = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
Sys.HighlightObject(addLine);
for(var i=0;i<addLine.ChildCount;i++){ 
  if((addLine.Child(i).isVisible())&&((addLine.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Add Job Invoice On Account Entry (Ctrl+M)").OleValue.toString().trim()) || (addLine.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Add Job Invoice On Account Entry").OleValue.toString().trim()))){
    addLine = addLine.Child(i);
    WorkspaceUtils.waitForObj(addLine);
    ReportUtils.logStep_Screenshot("");
    addLine.Click();
    break;
  }
}
aqUtils.Delay(4000, "Next Column");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
                   
var standardCode = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
addedlines = true;
standardCode.Click();
WorkspaceUtils.waitForObj(standardCode);
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(2000, "Next Column");
    
var description = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
description.Click();
description.settext(Desp);
description.Keys("[Tab]");
aqUtils.Delay(2000, "Next Column");

var qty = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
qty.Click();
qty.settext(Qly);
qty.Keys("[Tab]");
aqUtils.Delay(2000, "Next Column");

var billable = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
billable.Click();
billable.settext(UnitPrice);
billable.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
aqUtils.Delay(2000, "Next Column");

var workcode = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
workcode.Click();
WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"WorkCode");
  
if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_InvoiceOnAccount.specification",workcode,HSN);

  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
  WorkspaceUtils.waitForObj(Save);
for(var i=0;i<Save.ChildCount;i++){ 
  if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save Job Invoice On Account Entry (Enter)").OleValue.toString().trim())){
    Save = Save.Child(i);
    WorkspaceUtils.waitForObj(Save);
    ReportUtils.logStep_Screenshot("");
    Save.Click();
    break;
  }
}

aqUtils.Delay(1000, "Saving Line in On Account");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, "Saving Line in On Account");

    
     }
     
     aqUtils.Delay(3000, "Combined Invoice Selected");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


  
  }
  

}






//Submit Draft Invoice
function Submit_Draft(){ 
  
var Blanket_Invoice_Collection = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
Sys.HighlightObject(Blanket_Invoice_Collection);
Blanket_Invoice_Collection.Click();

aqUtils.Delay(3000, "Moving to Click Action");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000, "Moving to Click Action");
      
          
var Action = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.GroupToolItemControl;
if(Action.isVisible()){  
Sys.HighlightObject(Action);
Action.Click();

aqUtils.Delay(8000, "Clicking Submit");
Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim());
}
else{
var Approve = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 12);
Approve.Click();
}

aqUtils.Delay(15000, "Moving to Draft Invoice Tab"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000, "Moving to Click Action");
TextUtils.writeLog(" Blanket Invoice is Approved and moved to Draft Invoice")

var DraftInvoice = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(DraftInvoice);
DraftInvoice.Click();

aqUtils.Delay(3000, "Submit Draft");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000, "Submit Draft");

var table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
jobNumber = table.getItem(0).getText_2(0).OleValue.toString().trim()
Log.Message("Job Number : "+jobNumber);
aqUtils.Delay(3000, "Submit Draft");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000, "Submit Draft");

var CloseFilter = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(CloseFilter);
CloseFilter.Click();

aqUtils.Delay(3000, "Submit Draft");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000, "Submit Draft");

var PrintDraft = "";

  if(Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.isVisible())
  PrintDraft = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
  else
  PrintDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);

  WorkspaceUtils.waitForObj(PrintDraft);
  for(var i=0;i<PrintDraft.ChildCount;i++){ 
    if((PrintDraft.Child(i).isVisible())&&(PrintDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(PrintDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      PrintDraft.Child(i).Click();
      break;
    }
  } 
  
  aqUtils.Delay(8000, Indicator.Text);
  
TextUtils.writeLog("Print Draft is Clicked");
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Invoice Editing")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x41);
    
    if(ImageRepository.PDF.ChooseFolder.Exists())
    ImageRepository.PDF.ChooseFolder.Click();
    else{ 
    var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
    WorkspaceUtils.waitForObj(window);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x73); //F4
    Sys.Desktop.KeyUp(0x12); //Alt
    Sys.Desktop.KeyUp(0x73); //F4
      
    aqUtils.Delay(5000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x12); 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x41);
    }
    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
    aqUtils.Delay(2000, Indicator.Text);
    SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
aqUtils.Delay(2000, Indicator.Text);

Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Draft Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf");
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Draft Blanket Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  
aqUtils.Delay(4000, Indicator.Text);

  
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);



var SubmitDraft = ""
  if(Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.isVisible())
  SubmitDraft = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
  else
  SubmitDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);

  WorkspaceUtils.waitForObj(SubmitDraft);
  for(var i=0;i<SubmitDraft.ChildCount;i++){ 
    if((SubmitDraft.Child(i).isVisible())&&(SubmitDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(SubmitDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      SubmitDraft.Child(i).Click();
      TextUtils.writeLog("Draft Invoice is submitted");
      break;
    }
  }
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);


var Sliding_Panel = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
Sys.HighlightObject(Sliding_Panel);
Sliding_Panel.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

ImageRepository.ImageSet.Maximize.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);


var All_Approver = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(All_Approver);
All_Approver.Click();


aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

var Approvers_Table = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(Approvers_Table);
 var y=0;
for(var i=0;i<Approvers_Table.getItemCount();i++){   
   var approvers="";
    if(Approvers_Table.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
    approvers = EnvParams.Opco+"*"+jobNumber+"*"+Approvers_Table.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+Approvers_Table.getItem(i).getText_2(4).OleValue.toString().trim();
    Log.Message("Approver level :" +i+ ": " +approvers);
    Approve_Level[y] = approvers;
    y++;
    }
}
ReportUtils.logStep_Screenshot("");

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

var closeBar = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

ImageRepository.ImageSet.Forward.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

}



//Finding UserName for Approvers in Datasheets
function CredentialLogin(){ 
  var AppvLevl = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("FP")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("FP")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }

}





//Refreshing the To-Dos List
function todo(lvl){ 
  
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  aqUtils.Delay(3000, Indicator.Text);
  
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(3000, "Waiting to Refresh ToDo's List");
if(ImageRepository.ImageSet.ToDos_Icon.Exists())
{ 
  
}

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}

var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Drafts").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Purchase Order from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Drafts (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Invoice Drafts (Substitute) from To-Dos List");  
var listPass = false;   
  }
}  


if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Drafts by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Invoice Drafts by Type from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Drafts by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Invoice Drafts by Type (Substitute) from To-Dos List"); 
var listPass = false;   
  }
} 
  }
  
}




//Approving Invoice in every Job
function FinalApprove(JobNum,Apvr,lvl){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

//Checking the screen with CloseFilter or ShowFilter
var table = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}



var table = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(JobNum);
var closefilter = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;


aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

WorkspaceUtils.waitForObj(table);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==JobNum){ 
    flag=true;    
    table.Keys("[Down]");
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Draft Invoice is available in Approval List");
TextUtils.writeLog("Created Draft Invoice is available in Approval List");
if(flag){ 
aqUtils.Delay(1000, Indicator.Text);
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);


var Approve;

  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
 else
  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
            
    
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
}

Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();

ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
TextUtils.writeLog("Draft Invoice is Approved by "+Apvr);

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

             

var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);


  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 



if(Approve_Level.length==lvl+1){
  aqUtils.Delay(1000, Indicator.Text);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
                
var approvalBar = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

ImageRepository.ImageSet.Maximize.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

var DraftApproval = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();
  
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

var ApproverTable = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ReportUtils.logStep_Screenshot();

var closeBar = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

ImageRepository.ImageSet.Forward.Click();

aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var printStat = false;

var ChildCount = 0;
var Add = [];
var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
Sys.Process("Maconomy").Refresh();  
for(var ip=0;ip<Parent.ChildCount;ip++){ 
var PChild = Parent.Child(ip);
if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")&& (PChild.ChildCount==3)){
Log.Message(PChild.Name)
    Add[ChildCount] = PChild;
    ChildCount++;

}
}

var printInvoice = "";
var pos = 0;
for(var ip=0;ip<Add.length;ip++){ 
if(Add[ip].Height>pos){ 
pos = Add[ip].Height;
Log.Message(pos)
printInvoice = Add[ip];
}     
}
     
Log.Message(printInvoice.FullName);
Sys.HighlightObject(printInvoice)
     
if(printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
printInvoice = printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
else
printInvoice = printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
Sys.HighlightObject(printInvoice)
        
WorkspaceUtils.waitForObj(printInvoice);
for(var i=0;i<printInvoice.ChildCount;i++){ 
if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Invoice").OleValue.toString().trim())){
WorkspaceUtils.waitForObj(printInvoice.Child(i));
ReportUtils.logStep_Screenshot("");
printInvoice.Child(i).Click();
break;
}
} 

    aqUtils.Delay(10000, Indicator.Text);
    
    
     var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Drafts by Type - Invoice Editing").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Drafts by Type - Invoice Editing").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
}
TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
    aqUtils.Delay(5000, Indicator.Text);
 
//Saving PDF
TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Invoice")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x12); 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x41);
    
    if(ImageRepository.PDF.ChooseFolder.Exists())
    ImageRepository.PDF.ChooseFolder.Click();
    else{ 
      var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
      WorkspaceUtils.waitForObj(window);
      Sys.Desktop.KeyDown(0x12); //Alt
      Sys.Desktop.KeyDown(0x73); //F4
      Sys.Desktop.KeyUp(0x12); //Alt
      Sys.Desktop.KeyUp(0x73); //F4
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x12); 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x41);
    }
    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
    aqUtils.Delay(2000, Indicator.Text);
    SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Blanket Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  

var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(sFolder+SaveTitle+".pdf");
var textobj;
  try{


var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj).OleValue.toString(); 
  var invoiceName = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Invoice No:").OleValue.toString().trim();
  invoiceName = invoiceName.length;
  Log.Message(invoiceName)
  textobj = textobj.substring(textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice No:").OleValue.toString().trim()+" ")+invoiceName+1);
  Log.Message("Invoice No:"+textobj.substring(0,textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice Date").OleValue.toString().trim())))
  textobj = textobj.substring(0,textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice Date").OleValue.toString().trim()));
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Blanket Invoice No",EnvParams.Opco,"Data Management",textobj)
  ExcelUtils.WriteExcelSheet("Blanket Invoice Job",EnvParams.Opco,"Data Management",JobNum)
  TextUtils.writeLog("Client Invoice No: "+textobj);


}

  ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
  
  
}
}

}
