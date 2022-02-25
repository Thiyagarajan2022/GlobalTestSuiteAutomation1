//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT ObjectUtils
//USEUNIT ActionUtils


/**
 * This script create Quote and Client Approved Estimate for Main Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :02/10/2021
 * Modified Date(MM/DD/YYYY) : 02/23/2022
*/

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "InvoicePreparation";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var jobNumber,EmpNo = "";
var Estimatelines = [];
var IBudget_ID = "";
var IBudgetUnit = "";
var Descp = [];
var Hitpoint,Buss_Area_2 = "";

//Main Function
var Language = "";

function InvoicePlansOnAccount (){ 
  
TextUtils.writeLog("Create Invoice Plan - On Account Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

IBudget_ID = "";
IBudgetUnit = "";
Hitpoint,Buss_Area_2 = "";

  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  var invoicePreparation = ExcelUtils.getRowDatas("Invoice preparation Job",EnvParams.Opco);
  var AllocationWIP = ExcelUtils.getRowDatas("Job Invoice Allocation with WIP Job",EnvParams.Opco);
  var invoiceBudget = ExcelUtils.getRowDatas("Invoice from Budget Job",EnvParams.Opco);
  var invoiceAccount = ExcelUtils.getRowDatas("Invoice OnAccount Job",EnvParams.Opco);
  var writeoffInvoice = ExcelUtils.getRowDatas("Write Off Invoicing Job",EnvParams.Opco);
  
  template = ReadExcelSheet("Main Job Template",EnvParams.Opco,"Data Management");
  Log.Message((jobNumber!="")||(jobNumber!=null))
//  Log.Message(template.indexOf("FP")!=-1)
  Log.Message(invoicePreparation==jobNumber)
  Log.Message(AllocationWIP==jobNumber)
  Log.Message(invoiceBudget==jobNumber)
  Log.Message(invoiceAccount==jobNumber)
  Log.Message(writeoffInvoice==jobNumber)
//  if((jobNumber=="")||(jobNumber==null)){
  if(((jobNumber=="")||(jobNumber==null))||(invoicePreparation==jobNumber)||(AllocationWIP==jobNumber)||(invoiceBudget==jobNumber)||(invoiceAccount==jobNumber)||(writeoffInvoice==jobNumber)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  Log.Message(jobNumber);
  }
  if((invoicePreparation==jobNumber)||(AllocationWIP==jobNumber)||(invoiceBudget==jobNumber)||(invoiceAccount==jobNumber)||(writeoffInvoice==jobNumber)){
//    Log.Message(jobNumber+"Job Number is already used")
    jobNumber = "";
  }
  
  
  if((jobNumber=="")||(jobNumber==null)){ 
    //Creation of Job
    
    IBudget_ID = TestRunner.testCaseId;
    IBudgetUnit = TestRunner.unitName; 
    TestRunner.TempUnit = IBudgetUnit;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Invoice Plan - On Account")
    }
    
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Invoice Plan - On Account")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    jobNumber = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    
    if((jobNumber=="")||(jobNumber==null)){
      
    TestRunner.unitName = "JobCreation_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job");
    Log.PushLogFolder(FolderID);
    Runner.CallMethod("Dependency_Job_Creation.createJob",jobSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    jobNumber = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    
    //Creation of Budget
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var budgetSheet = ExcelUtils.getColumnDatas("Budget sheet",EnvParams.Opco)
    if(budgetSheet==""){ 
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Invoice Plan - On Account")
    }
    ExcelUtils.setExcelName(workBook, budgetSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Budget")
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var WE_Number = ExcelUtils.getRowDatas("Working Estimate_"+serialOder,EnvParams.Opco)
    if((WE_Number=="")||(WE_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;

    TestRunner.unitName = "CreateBudget_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job Budget");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job Budget");
    Log.PushLogFolder(FolderID);
    Runner.CallMethod("Dependency_JobBudget_Creation.createBudget",budgetSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
  //Creation of Quote 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var quoteSheet = ExcelUtils.getColumnDatas("Quote Sheet",EnvParams.Opco)
    if(quoteSheet==""){ 
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Invoice Plan - On Account")
    }
    ExcelUtils.setExcelName(workBook, quoteSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Quote")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var CE_Number = ExcelUtils.getRowDatas("Client Approved Estimate_"+serialOder,EnvParams.Opco)
    if((CE_Number=="")||(CE_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
 
    TestRunner.unitName = "CreateQuote_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Quote");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Quote");
    Log.PushLogFolder(FolderID);
    Runner.CallMethod("Dependency_JobQuote_Creation.CreateQuote",quoteSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
TestRunner.testCaseId = IBudget_ID;
TestRunner.unitName = IBudgetUnit;

}


Log.Message(jobNumber)
Log.Message(template)


TestRunner.JiraStat = true;
TestRunner.JiraUpdate = true;




ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "InvoicePreparation";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
jobNumber,EmpNo = "";
Estimatelines = [];
Descp = [];

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "InvoicingPlans-Invoice On Account started::"+STIME);
getDetails();
gotoMenu();
gotoInvoicing();
invoiceOnAccount();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);
//for(var i=level;i<ApproveInfo.length;i++){
//level=i;
//WorkspaceUtils.closeMaconomy();
//aqUtils.Delay(10000, Indicator.Text);
//var temp = ApproveInfo[i].split("*");
//Restart.login(temp[2]);
for(var i=level;i<ApproveInfo.length;i++){
  level = i;
  
  

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);

Maconomy_Index = WorkspaceUtils.Maconomy_Parent;
Maconomy_Index++;
WorkspaceUtils.Maconomy_Parent = Maconomy_Index;
Log.Message(Maconomy_Index);


// Restarting maconomy with Approver Logins
var temp = ApproveInfo[i].split("*");
var Macscreen = WorkspaceUtils.switch_Maconomy(temp[2])
if(Macscreen=="Screen Not Found"){
Restart.login(temp[2]);
}
aqUtils.Delay(5000, Indicator.Text);

Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(temp[2]);
//Refreshing To-Do's List to find Submitted Job
aqUtils.Delay(5000, Indicator.Text);

aqUtils.Delay(5000, Indicator.Text);
//todo(temp[3]);
ActionUtils.ToDos_Selection(Maconomy_ParentAddress, i, temp[3] ,"Approve Invoice Drafts" ,"Approve Invoice Drafts by Type" , "Approve Invoice Drafts (Substitute)" , "Approve Invoice Drafts by Type (Substitute)")


FinalApprove(temp[1],temp[2],i);


}
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(5000, Indicator.Text);
}


function getDetails(){ 
sheetName ="InvoicePreparation";  

  ExcelUtils.setExcelName(workBook, sheetName, true);
  EmpNo = ExcelUtils.getColumnDatas("Employee Number",EnvParams.Opco)

  if((EmpNo=="")||(EmpNo==null))
  ValidationUtils.verify(false,true,"Employee Number is needed for Transfer Invoice plan - on Account");
  
  sheetName ="InvoicePreparation";  


  

  ExcelUtils.setExcelName(workBook, sheetName, true);
  percentage = ExcelUtils.getColumnDatas("Percentage",EnvParams.Opco)
  if((percentage=="")||(percentage==null))
  ValidationUtils.verify(false,true,"Percentage is needed for Invoice Preparation");
  

  ExcelUtils.setExcelName(workBook, sheetName, true);
  Hitpoint = ExcelUtils.getColumnDatas("Sent To Hitpoint",EnvParams.Opco)
  if(Hitpoint.toUpperCase()=="YES"){
  Buss_Area_2 = ExcelUtils.getColumnDatas("Business Area 2",EnvParams.Opco);
  }

  
}

function gotoMenu(){ 

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_Jobs_from_workspace(); //Select Jobs Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());


}

function gotoInvoicing(){ 
waitUntil_MaconomyScreen_loaded_Completely();

waitUntil_MaconomyScreen_loaded_Completely();

  var allJobs = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();


waitUntil_MaconomyScreen_loaded_Completely();

waitUntil_MaconomyScreen_loaded_Completely();


  var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var firstcell = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  var closeFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);


TemplateJob = ""
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
      TemplateJob = table.getItem(v).getText_2(4).OleValue.toString().trim()
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to for Direct Invocing");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy for Direct Invocing"); 
  closeFilter.Click();

  aqUtils.Delay(1000, Indicator.Text);
  waitUntil_MaconomyScreen_loaded_Completely();
  
//  var clientApproved = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  var clientApproved = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget", "2",1);
  clientApproved = clientApproved.SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 5)
  WorkspaceUtils.waitForObj(clientApproved);
  if(clientApproved.background!=10674625){
    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
    }
    
//  var workingEstimate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
  var workingEstimate = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget", "2",1);
  workingEstimate = workingEstimate.SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 5)
  WorkspaceUtils.waitForObj(workingEstimate);
  if(workingEstimate.background!=10674625){
    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
    }
      
//  var lastInvoice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
//  var totalInvoice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
//  var billingPrice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
//  var netInvoiceOnAcc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  
  var Budgeting = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Budgeting");
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  
  aqUtils.Delay(100, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
  var Estimate = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget", "1",9)
  Estimate.Keys("Client Approved Estimate");
  aqUtils.Delay(100, Indicator.Text);
  
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
    var FullBudget = ActionUtils.getObjectAddress_JavaClasssName_Index_withTabText(Maconomy_ParentAddress,"TabControl", "6","Full Budget");
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var BudgetGrid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
  WorkspaceUtils.waitForObj(BudgetGrid);
  var ii=0;
  
  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")||(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")){ 
         if(EnvParams.Country.toUpperCase()=="INDIA")
         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(12).OleValue.toString().trim();
         else
         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim();
         Log.Message(Estimatelines[ii]);
         ii++;
    }
  }

var info = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Home");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
info.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var workCodeAdd = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Work Codes");
WorkspaceUtils.waitForObj(workCodeAdd);
workCodeAdd.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

workCodeList = [];
for(var i=0;i<Estimatelines.length;i++){

var temp = Estimatelines[i].split("*");
if(temp[0]!=""){
workCodeList[i] = temp[0];
Log.Message(workCodeList[i])
}
}

workActivity = [];
var i=0
var WorkCodeGrid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
for(var v=0;v<WorkCodeGrid.getItemCount();v++){ 
  for(var y=0;y<workCodeList.length;y++){ 
  if(WorkCodeGrid.getItem(v).getText(0).OleValue.toString().trim()==workCodeList[y]){ 
    workActivity[i] = WorkCodeGrid.getItem(v).getText(0).OleValue.toString().trim()+"*"+WorkCodeGrid.getItem(v).getChecked(14)
    Log.Message(workActivity[i]);
    i++;
  }
  
  }
}

  var Invoicing = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Invoicing");
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

  var Plan = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Plan");
//  var Plan = Aliases.Maconomy.DirectInvoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Plan);
  Plan.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var spec_Overview = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Specification, Overview");
  WorkspaceUtils.waitForObj(spec_Overview);
  spec_Overview.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

for(var i = 0;i<Estimatelines.length;i++){ 
  var temp_1 = Estimatelines[i].split("*");
for(var j = 0;j<workActivity.length;j++){ 
  var temp_2 = workActivity[j].split("*");
  
  if((temp_1[0]==temp_2[0])&&(temp_2[1]=="true")){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var AddLines = ActionUtils.getObjectAddress_JavaClasssName_Index_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl",4,"Add Invoicing Plan Line")
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
AddLines.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
var Date = table.SWTObject("McDatePickerWidget", "");
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
Date.Click();
Date.setText(aqDateTime.Today())
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
var Descp = table.SWTObject("McTextWidget", "", 2);
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Descp.Click();
Descp.setText(temp_1[1]);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var Qty = table.SWTObject("McTextWidget", "", 3);
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
Qty.Click();
Qty.setText(temp_1[2]);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var UnitPrice = table.SWTObject("McTextWidget", "", 3);
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
UnitPrice.Click();
UnitPrice.setText(temp_1[3]);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var workCode = table.SWTObject("McValuePickerWidget", "", 5);
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
workCode.Click();
WorkspaceUtils.SearchByValue(workCode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),temp_1[0],"WorkCode");
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

if(temp_1[0].indexOf("T")==0){
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var Employee = table.SWTObject("McValuePickerWidget", "", 5);
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
Employee.Click();
WorkspaceUtils.SearchByValue(Employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
aqUtils.Delay(1000, Indicator.Text);

Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyUp(0x10);
aqUtils.Delay(1000, Indicator.Text);

Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyUp(0x10);
aqUtils.Delay(1000, Indicator.Text);

Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyUp(0x10);
aqUtils.Delay(1000, Indicator.Text);

Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyUp(0x10);
aqUtils.Delay(1000, Indicator.Text);

Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyUp(0x10);
aqUtils.Delay(1000, Indicator.Text);

Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyUp(0x10);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
}
var selectManually = table.SWTObject("McPlainCheckboxView", "", 6).SWTObject("Button", "")
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
selectManually.Click();
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var approve = table.SWTObject("McPlainCheckboxView", "", 6).SWTObject("Button", "");
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
approve.Click();
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var save = ActionUtils.getObjectAddress_JavaClasssName_Index_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl",2,"Save Invoicing Plan Line");
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
save.Click();
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ValidationUtils.verify(true,true,temp_1[0]+" is added for Invoice plan")
TextUtils.writeLog(temp_1[0]+" is added for Invoice plan");
  break;
  }
  
}


}

//var AddLines = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//var Date = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
//var Descp = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//var Qty = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
//var UnitPrice = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
//var workCode = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//var selectManually = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
//var approve = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
//var save = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;


//var TobePlanned_Job = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
//var TobePlanned_Invocie = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget2;

var groupWidget = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"McGroupWidget",1,"Composite",1);
var TobePlanned_Job = groupWidget.SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
var TobePlanned_Invocie = groupWidget.SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 3);

if((TobePlanned_Job.getText()=="0.00")&&(TobePlanned_Invocie.getText()=="0.00")){ 
  ValidationUtils.verify(true,true,"Invoice Plan is balanced")
}
else{ 
 ValidationUtils.verify(false,true,"Invoice Plan is Not balanced") 
}

  }
}


function invoiceOnAccount(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var InvoiceAccount = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Invoice On Account");
InvoiceAccount.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

//var TransferPlan = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.SingleToolItemControl;
var TransferPlan = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"GroupToolItemControl","More Actions");
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl2;
TransferPlan.Click();
aqUtils.Delay(5000, "Transfer Invoicing Plan");
TransferPlan.PopupMenu.Click("Transfer Invoicing Plan");

/*
var SelectionBilling = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

  for(var t=0;t<SelectionBilling.getItemCount()-1;t++){ 
      for(var g = 0;g<Estimatelines.length;g++){
        var temp = Estimatelines[g].split("*");
        if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])==0){
          if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf("T")==0){
      aqUtils.Delay(100, Indicator.Text);
      SelectionBilling.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
      var Employee = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
      WorkspaceUtils.waitForObj(Employee);
      Employee.Click();
//      if((Employee.getText()=="")||(Employee.getText()==null)){ 
      if((EmpNo!="")&&(EmpNo!=null)){
      Employee.HoverMouse();
      Employee.Click();
      WorkspaceUtils.SearchByValue(Employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
      
var SaveStat = true;
      var Save = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite;
      Log.Message(Save.FullName)
      WorkspaceUtils.waitForObj(Save);
      for(var i=0;i<Save.ChildCount;i++){
        Log.Message(Save.Child(i).Name)
        Log.Message(Save.Child(i).toolTipText)
        if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Save Invoice Selection Line").OleValue.toString().trim())){
          Save = Save.Child(i);
          WorkspaceUtils.waitForObj(Save);
          ReportUtils.logStep_Screenshot("");
          Save.Click();
          SaveStat = false;
          break;
        }
        
      } 
      
      }
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }else{ 
    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
    }

      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(1000, Indicator.Text);
      break;
    }
    }
    }
    if(i<SelectionBilling.getItemCount()-2){
    SelectionBilling.Keys("[Down]");
    }
    }













var Action = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.GroupToolItemControl;
WorkspaceUtils.waitForObj(Action);
ReportUtils.logStep_Screenshot("");
Action.Click();
aqUtils.Delay(2000, Indicator.Text);;
//if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//}else{ 
//ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//}
Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transfer Invoicing Plan").OleValue.toString().trim());

*/

ReportUtils.logStep_Screenshot("");
aqUtils.Delay(100, "Transfer Invoicing Plan");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
aqUtils.Delay(8000, "Transfer Invoicing Plan");
var plan = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "Transfer Invoicing Plan").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Transfer Invoicing Plan");
//Aliases.Maconomy.Transfer_Invoicing_Plan.Composite.Composite.Composite.Composite.SWTObject("Button", "Transfer Invoicing Plan");
plan.Click();
aqUtils.Delay(8000, "Transfer Invoicing Plan");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
aqUtils.Delay(100, "Billing Price");
var BudgetAmount = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"McGroupWidget",1,"Composite",1);
BudgetAmount = BudgetAmount.SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
BudgetAmount = BudgetAmount.getText();
Log.Message(BudgetAmount);
var Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve");
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 5);
WorkspaceUtils.waitForObj(Approve);
ReportUtils.logStep_Screenshot("");
Approve.Click();
TextUtils.writeLog("Approve is clicked");
aqUtils.Delay(1000, "Approve is Clicked");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

TextUtils.writeLog("Approve is clicked");
aqUtils.Delay(1000, "Approve is Clicked");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
var SelectionBilling = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
//Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Descp = [];
  var Ds = 0;
  for(var t=0;t<SelectionBilling.getItemCount();t++){ 
   Descp[Ds] = SelectionBilling.getItem(t).getText_2(10).OleValue.toString().trim()+"*"+SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim();
   Log.Message(Descp[Ds])
   Ds++; 
  }

var DraftInvoice = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Draft Invoices");
//Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(DraftInvoice);
ReportUtils.logStep_Screenshot("");
DraftInvoice.Click();
aqUtils.Delay(1000, Indicator.Text);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

var draftNo = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
draftNo = draftNo.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(draftNo);
draftNo.Keys("[Tab][Tab][Tab]");
var billiablePrice = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
billiablePrice = billiablePrice.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(billiablePrice);
billiablePrice.Click();
billiablePrice.setText(BudgetAmount);

//var DraftTable = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var flag=true;
//  for(var v=0;v<DraftTable.getItemCount();v++){ 
//    if(DraftTable.getItem(v).getText_2(3).OleValue.toString().trim()==BudgetAmount){ 
//      flag=true;
//      break;
//    }
//    else{ 
//      DraftTable.Keys("[Down]");
//    }
//  }
 ValidationUtils.verify(true,flag,"Invoice is available to submit Draft")
  if(flag){
var CloseFilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
ReportUtils.logStep_Screenshot("");
CloseFilter.Click();
aqUtils.Delay(1000, Indicator.Text);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
  var DraftEditing = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Draft Editing");
  DraftEditing.Click();
  var grid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid","2");
  var Save = ActionUtils.getObjectAddress_JavaClasssName_Index_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl",3,"Save Invoice Line");
//  grid.Keys("[Tab]");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  for(var i=0;i<grid.getItemCount()-1;i++){ 
//    Log.Message(grid.getItem(i).getText_2(1).OleValue.toString().trim())
    for(var j=0;j<Descp.length;j++){ 
      var temp = Descp[j].split("*")
//      Log.Message(grid.getItem(i).getText_2(1).OleValue.toString().trim())
//      Log.Message(temp[0])
      if(grid.getItem(i).getText_2(1).OleValue.toString().trim()==temp[0]){ 
        if(grid.getItem(i).getText_2(1).OleValue.toString().trim()!=temp[1]){
          grid.Keys("[Tab]");
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
              }
          var Des = grid.SWTObject("McTextWidget", "", 2);
          Des.Click();
          Des.setText(temp[1]);
          aqUtils.Delay(1000, Indicator.Text);
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
            }
            
          Save = ActionUtils.getObjectAddress_JavaClasssName_Index_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl",3,"Save Invoice Line");
          Log.Message(Save.FullName);
          Save.Click();
          aqUtils.Delay(1000, Indicator.Text);
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
              }
          Sys.Desktop.KeyDown(0x10);
          Sys.Desktop.KeyDown(0x09);
          aqUtils.Delay(1000, Indicator.Text);
          Sys.Desktop.KeyUp(0x10);
          Sys.Desktop.KeyUp(0x09);
          aqUtils.Delay(1000, Indicator.Text);
          break;
          }
      }
    }
  var grid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid","2");
  Sys.HighlightObject(grid)
  Log.Message(i)
  Log.Message(grid.getItemCount()-2)
  Log.Message(i<grid.getItemCount()-2)
  if(i<grid.getItemCount()-2){
  grid.Keys("[Down]");
  }
  }  
  
  
var SubmitDraft = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Submit Draft");
Sys.HighlightObject(SubmitDraft);
SubmitDraft.Click();


  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
   
//var Excl_Tax = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//var grandTotal = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);

var groupWidget = getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"McGroupWidget",2,"Composite",2);
var Excl_Tax = groupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
var grandTotal = groupWidget.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);


//Finding Payment Terms
var break_MainLoop = false;
//var ParentAdd = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 6)

//var Payment_Terms = "";
//for(var i=0;i<ParentAdd.ChildCount;i++){ 
//  var temp = ParentAdd.Child(i);
//  for(var j=0;j<temp.ChildCount;j++){ 
//    if(temp.Child(j).Name.indexOf("McPopupPickerWidget")!=-1){
//      Payment_Terms = temp.Child(j);
//      break_MainLoop = true;
//      break;
//    }
//  }
//  
//  if(break_MainLoop){ 
//    break;
//  }
//}


//var Payment_Terms =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").
//SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).
//SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2)

var groupWidget = getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"McGroupWidget",1,"Composite",1);
var Payment_Terms = groupWidget.SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2)

Excl_Tax = Excl_Tax.getText().OleValue.toString().trim();
grandTotal = grandTotal.getText().OleValue.toString().trim();
Payment_Terms = Payment_Terms.getText().OleValue.toString().trim();
Payment_Terms = Payment_Terms.replace(/[^0-9]+/g, "");;
var Q_total = 0;
//var specification = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var specification = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
  var q = 0;
QuoteDetails = [];
var InvoiceMPL = "InvoiceMPL";
for(var i=0;i<specification.getItemCount();i++){ 

  var Q_Desp = specification.getItem(i).getText_2(1).OleValue.toString().trim();
  if(Q_Desp!=""){
    
  var Q_Qty = specification.getItem(i).getText_2(2).OleValue.toString().trim();
  var Q_Billing = specification.getItem(i).getText_2(3).OleValue.toString().trim();
  var Q_BillingTotal = specification.getItem(i).getText_2(4).OleValue.toString().trim();
  var Q_Tax1 = specification.getItem(i).getText_2(7).OleValue.toString().trim();
  var Q_Tax2 = specification.getItem(i).getText_2(9).OleValue.toString().trim();
  var Q_Tax1currency = specification.getItem(i).getText_2(8).OleValue.toString().trim();
  var Q_Tax2currency = specification.getItem(i).getText_2(10).OleValue.toString().trim();
//  var Q_total = parseFloat(Q_BillingTotal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
//  QuoteDetails[q] = Q_Desp+"*"+Q_Qty+"*"+Q_Billing+"*"+Q_BillingTotal+"*"+Q_Tax1+"*"+Q_Tax2+"*"+Q_Tax1currency+"*"+Q_Tax2currency+"*"+Q_total.toFixed(2)+"*";
//  Log.Message(QuoteDetails[q]);
  Q_total =parseFloat(Q_BillingTotal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
  Log.Message(Q_total);
  QuoteDetails[q] = Q_Desp+"*"+Q_Qty+"*"+Q_Billing+"*"+Q_BillingTotal+"*"+Q_Tax1+"*"+Q_Tax2+"*"+Q_Tax1currency+"*"+Q_Tax2currency+"*"+Q_total.toFixed(2)+"*";
  Log.Message(QuoteDetails[q]);
  q++;
  ExcelUtils.setExcelName(workBook,InvoiceMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,InvoiceMPL,Q_Desp);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,InvoiceMPL,Q_Qty);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,InvoiceMPL,Q_Billing);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TotalBilling_"+q,InvoiceMPL,Q_BillingTotal);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1_"+q,InvoiceMPL,Q_Tax1);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2_"+q,InvoiceMPL,Q_Tax2);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1currency_"+q,InvoiceMPL,Q_Tax1currency);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2currency_"+q,InvoiceMPL,Q_Tax2currency);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Total_"+q,InvoiceMPL,Q_total);

  }
  }

  ExcelUtils.setExcelName(workBook,InvoiceMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TOTAL EXC. TAX",InvoiceMPL,Excl_Tax);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Invoice TOTAL",InvoiceMPL,grandTotal);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Payment Terms",InvoiceMPL,Payment_Terms);
  
  
  aqUtils.Delay(2000, Indicator.Text);
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  var PrintDraft;
   PrintDraft = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Print Draft");
   Sys.HighlightObject(PrintDraft);
   PrintDraft.Click();
  

    aqUtils.Delay(4000, Indicator.Text);
    aqUtils.Delay(4000, Indicator.Text);
    WorkspaceUtils.savePDF_localDirectory("PDF Draft Invoice","Print Invoice Editing");
     

var appvBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",2)
  appvBar = appvBar.SWTObject("TabControl", "");
appvBar.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
ImageRepository.ImageSet.Maximize.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var DraftApproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","All Approval Actions");
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var ApproverTable = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
WorkspaceUtils.waitForObj(ApproverTable);
 var y=0;
for(var i=0;i<ApproverTable.getItemCount();i++){   
   var approvers="";
    if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
    approvers = EnvParams.Opco+"*"+jobNumber+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
    Log.Message("Approver level :" +i+ ": " +approvers);
    Approve_Level[y] = approvers;
    y++;
    }
}
ReportUtils.logStep_Screenshot("");
var closeBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
Log.Message(OpCo2[2]);
Log.Message(Project_manager);
if(OpCo2[2]==Project_manager){
level = 1;

var Approve;
////var Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
//  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.isVisible())
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
// else
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
            
var Approve;
Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve Draft");
Sys.HighlightObject(Approve);
Approve.Click();

ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved Draft Invoice");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

  
}

}

}

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
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
    Log.Message(temp);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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
//WorkspaceUtils.closeAllWorkspaces();
}

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
  
}else{ 
  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
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
if((temp.indexOf("Approve Invoice Drafts (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf("Approve Invoice Drafts Substitute) (")!=-1)&&(temp1.length==3)){ 
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
if((temp.indexOf("Approve Invoice Drafts by Type (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf("Approve Invoice Drafts by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Invoice Drafts by Type (Substitute) from To-Dos List"); 
var listPass = false;   
  }
} 
  }
  
}




function FinalApprove(JobNum,Apvr,lvl){ 
  aqUtils.Delay(4000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
                                       
//var table = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder;
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "");
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

//var table = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
//var firstCell = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
var firstCell = table.SWTObject("McValuePickerWidget", "");
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(JobNum);
//var closefilter = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
var closefilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
                                              

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
WorkspaceUtils.waitForObj(table);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==JobNum){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Draft Invoice is available in Approval List");
TextUtils.writeLog("Created Draft Invoice is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
var Approve;
Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve Draft")


//WorkspaceUtils.waitForObj(Approve);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();

ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
TextUtils.writeLog("Draft Invoice is Approved by "+Apvr);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

              
loginPer = eval(Maconomy_ParentAddress).WndCaption;
loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 

  



if(Approve_Level.length==lvl+1){
  aqUtils.Delay(1000, Indicator.Text);

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
//var printInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
                  
//var approvalBar = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
var approvalBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",2)
approvalBar = approvalBar.SWTObject("TabControl", "");
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
    ImageRepository.ImageSet.Maximize.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//var DraftApproval = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
var DraftApproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","All Approval Actions");
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();
  

//var ApproverTable = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var ApproverTable = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
WorkspaceUtils.waitForObj(ApproverTable);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
ReportUtils.logStep_Screenshot();

//var closeBar = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
//var closeBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",1,"Composite",2)
//closeBar = closeBar.SWTObject("TabControl", "");
var closeBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
ImageRepository.ImageSet.Forward.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


var printStat = false;

if(Hitpoint.toUpperCase()!="YES"){
  

  var printInvoice = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Print Invoice");
  Sys.HighlightObject(printInvoice);
  printInvoice.Click();
  
    aqUtils.Delay(10000, Indicator.Text);
    
    

    
     var p = eval(WorkspaceUtils.Sys_Maconomy_Parent);
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
  
var PDF_Location = ""

PDF_Location = WorkspaceUtils.savePDF_localDirectory("PDF Invoice","Print Job Invoice");

  


var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(PDF_Location);
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
  ExcelUtils.WriteExcelSheet("Invoice preparation No",EnvParams.Opco,"Data Management",textobj)
  ExcelUtils.WriteExcelSheet("Invoice preparation Job",EnvParams.Opco,"Data Management",PONum)
  ExcelUtils.WriteExcelSheet("Client Invoice No",EnvParams.Opco,"Data Management",textobj)
  TextUtils.writeLog("Client Invoice No: "+textobj);

}else{ 
  aqUtils.Delay(2000, Indicator.Text);
//  var Sent_To_Hitpoint = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SWTObject("SingleToolItemControl", "", 15);
  var Sent_To_Hitpoint = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Send To Hitpoint");
  Sys.HighlightObject(Sent_To_Hitpoint);
  Sent_To_Hitpoint.Click();
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  var menuBar = eval(Maconomy_ParentAddress).SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  WorkspaceUtils.closeAllWorkspaces();
    aqUtils.Delay(5000, Indicator.Text);
  goto_Hitpoint_Billing();
  Check_Hitpoint_Status();
  
  
}


  ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
  }
  
}
}

}




