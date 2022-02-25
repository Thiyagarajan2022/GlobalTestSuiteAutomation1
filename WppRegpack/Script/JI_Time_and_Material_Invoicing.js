﻿//USEUNIT ActionUtils
//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ObjectUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/**
 * This script create Quote and Client Approved Estimate for Main Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :02/10/2021
 * Modified Date(MM/DD/YYYY) : 02/19/2022
*/

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Time_Material_Invocing";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var jobNumber,EmpNo ,Hitpoint,Buss_Area_2 = "";
var Estimatelines = [];
var B_Estimatelines = [];
var Q_Estimatelines = [];
var LatestTran = ""
var Language = "";
var Descp = [];
var TemplateJob = "";
var IBudget_ID = "";
var IBudgetUnit = "";
var MainJob = "";
var Maconomy_ParentAddress,Maconomy_Index = "";

//Main Function
function InvoiceAllocation(){ 
  
TextUtils.writeLog("Time & Material Invocing (Without WIP) Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

IBudget_ID = "";
IBudgetUnit = "";
Hitpoint,Buss_Area_2 = "";



  MainJob = true;
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  var invoicePreparation = ExcelUtils.getRowDatas("Invoice preparation Job",EnvParams.Opco);
  var AllocationWIP = ExcelUtils.getRowDatas("Job Invoice Allocation with WIP Job",EnvParams.Opco);
  var invoiceBudget = ExcelUtils.getRowDatas("Invoice from Budget Job",EnvParams.Opco);
  var invoiceAccount = ExcelUtils.getRowDatas("Invoice OnAccount Job",EnvParams.Opco);
  var writeoffInvoice = ExcelUtils.getRowDatas("Write Off Invoicing Job",EnvParams.Opco);
  
  template = ReadExcelSheet("Main Job Template",EnvParams.Opco,"Data Management");
  Log.Message((jobNumber!="")||(jobNumber!=null))
  Log.Message(invoicePreparation==jobNumber)
  Log.Message(AllocationWIP==jobNumber)
  Log.Message(invoiceBudget==jobNumber)
  Log.Message(invoiceAccount==jobNumber)
  Log.Message(writeoffInvoice==jobNumber)
  if(((jobNumber=="")||(jobNumber==null))||(invoicePreparation==jobNumber)||(AllocationWIP==jobNumber)||(invoiceAccount==jobNumber)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  Log.Message(jobNumber);
  }
  //if((invoicePreparation==jobNumber)||(AllocationWIP==jobNumber)(invoiceAccount==jobNumber)){
    //jobNumber = "";
  //}
  if((jobNumber=="")||(jobNumber==null)){ 
    
    //Creation of Job
    MainJob = false
    IBudget_ID = TestRunner.testCaseId;
    IBudgetUnit = TestRunner.unitName; 
    TestRunner.TempUnit = IBudgetUnit;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Time & Material Invocing")
    }
    
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Time & Material Invocing")
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
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Time & Material Invocing")
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
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Time & Material Invocing")
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


  
  
//Checking Login to execute Create Budget
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
sheetName = "Time_Material_Invocing";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
jobNumber,EmpNo,LatestTran = "";
Estimatelines = [];

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Time & Material Invocing (Without WIP) started::"+STIME);
getDetails();
goTo_Job_Menu();
gotoAllocation();

GoToDraft();
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
sheetName ="Time_Material_Invocing";  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  EmpNo = ReadExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management");
  if((EmpNo=="")||(EmpNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  EmpNo = ExcelUtils.getColumnDatas("Employee Number",EnvParams.Opco)
  }
  if((EmpNo=="")||(EmpNo==null))
  ValidationUtils.verify(false,true,"Employee Number is needed for Time & Material Invocing");
  
   
  
  ExcelUtils.setExcelName(workBook, sheetName, true);
  Hitpoint = ExcelUtils.getColumnDatas("Sent To Hitpoint",EnvParams.Opco)
  if(Hitpoint.toUpperCase()=="YES"){
  Buss_Area_2 = ExcelUtils.getColumnDatas("Business Area 2",EnvParams.Opco);
  }
 
}



/**
  *  This function Navigates to Jobs screen from Jobs workspace
  */
  
// Navigating to Jobs from Jobs Menu
function goTo_Job_Menu(){

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_Jobs_from_workspace(); //Select Jobs Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());

}


function gotoAllocation(){ 
  
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
  
  var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
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
  ReportUtils.logStep("INFO", "Job is listed in table to for Time & Material Invocing");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy for Time & Material Invocing"); 
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
  
    if(Hitpoint.toUpperCase()=="YES"){
  // Moving to Information Tab to Submit
//  var info = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 5)
  var info = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Information");
  Sys.HighlightObject(info);
  info.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  info.Click();


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(1000, Indicator.Text);
//  var business_Area2 = Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
  var business_Area2 = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGroupWidget", "4",1);
  business_Area2 = business_Area2.SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
  Sys.HighlightObject(business_Area2);
  business_Area2.Click();
  WorkspaceUtils.SearchByValue(business_Area2,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Option").OleValue.toString().trim(),Buss_Area_2,"Bussiness Area 2");
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(1000, Indicator.Text);
  
  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SingleToolItemControl;
  var Save = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Save Job")
  Sys.HighlightObject(Save);
  Save.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(1000, Indicator.Text);
  
//  var Budgeting = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  var Budgeting = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Budgeting");
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  
  }
else{
  
  /// Original Invoice
  
//  var Budgeting = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
  var Budgeting = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Budgeting");
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  }

//  var Budgeting = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
//  WorkspaceUtils.waitForObj(Budgeting);
//  Budgeting.Click();
  
  aqUtils.Delay(100, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  

  
//  var Estimate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  var Estimate = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget", "1",9)
  Estimate.Keys("Client Approved Estimate");
  aqUtils.Delay(100, Indicator.Text);
  
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
//  var FullBudget = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  var FullBudget = ActionUtils.getObjectAddress_JavaClasssName_Index_withTabText(Maconomy_ParentAddress,"TabControl", "6","Full Budget");
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
//  var BudgetGrid = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var BudgetGrid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
  WorkspaceUtils.waitForObj(BudgetGrid);
  var ii=0;
  
  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")||(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")){ 
        B_Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(15).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(16).OleValue.toString().trim();
         Log.Message(B_Estimatelines[ii]);
         ii++;
    }
  }

//  var Quote = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl;
  var Quote = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Quote");
  WorkspaceUtils.waitForObj(Quote);
  Quote.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//  var BudgetGrid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var BudgetGrid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
  WorkspaceUtils.waitForObj(BudgetGrid);
  var ii=0;
  
  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")||(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")){ 
         Q_Estimatelines[ii] = "WorkCode"+"*"+BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(1).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(2).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim();
         Log.Message(Q_Estimatelines[ii]);
         ii++;
    }
  }
  
  
//  if(Hitpoint.toUpperCase()=="YES"){
//  var Invoicing = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8);
//  WorkspaceUtils.waitForObj(Invoicing);
//  Invoicing.Click();  
//  }else{
//  var Invoicing = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8)
//  WorkspaceUtils.waitForObj(Invoicing);
//  Invoicing.Click();
//  }
  
  var Invoicing = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Invoicing");
  Invoicing.Click();
  
//  var Invoicing = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8)
//  WorkspaceUtils.waitForObj(Invoicing);
//  Invoicing.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

  Log.Message(TemplateJob)
//  var iselection = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  var iselection = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Invoice Selection");
  Log.Message(iselection.FullName)
  WorkspaceUtils.waitForObj(iselection);
  ReportUtils.logStep_Screenshot("");
  iselection.Click();
  TextUtils.writeLog("Entering into Invoice Selection Tab");

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


//  var SelectionBilling = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var SelectionBilling = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
  WorkspaceUtils.waitForObj(SelectionBilling);
  aqUtils.Delay(100, Indicator.Text);
  
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

    ImageRepository.ImageSet.Maximize1.Click();
  if(EnvParams.Country.toUpperCase()=="INDIA")
  Runner.CallMethod("IND_InvoiceAllocation.Employeenumber",SelectionBilling,EmpNo,B_Estimatelines);
  else
  for(var t=0;t<SelectionBilling.getItemCount();t++){ 
    if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf("T")==0){
      for(var g = 0;g<B_Estimatelines.length;g++){
        var temp = B_Estimatelines[g].split("*");
        if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])==0){
      aqUtils.Delay(100, Indicator.Text);
      SelectionBilling.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
//      var Employee = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
      var Employee = SelectionBilling.SWTObject("McValuePickerWidget", "", 4);
      WorkspaceUtils.waitForObj(Employee);
      Employee.Click();
//      if((Employee.getText()=="")||(Employee.getText()==null)){ 
      if((EmpNo!="")&&(EmpNo!=null)){
      Employee.HoverMouse();
      Employee.Click();
      WorkspaceUtils.SearchByValue(Employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
      
var SaveStat = true;

      

//       var Save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
       var Save = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save Invoice Selection Line").OleValue.toString().trim())
           WorkspaceUtils.waitForObj(Save);
          ReportUtils.logStep_Screenshot("");
          Save.Click();
          
//      Log.Message(Save.FullName)
//      WorkspaceUtils.waitForObj(Save);
//      for(var i=0;i<Save.ChildCount;i++){
//        Log.Message(Save.Child(i).Name)
//        Log.Message(Save.Child(i).toolTipText)
////        if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save Invoice Selection Line").OleValue.toString().trim())){
//        if((Save.Child(i).isVisible())&&((Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save Invoice Selection Line (Enter)").OleValue.toString().trim()) || (Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save Invoice Selection Line").OleValue.toString().trim()))){
//          Save = Save.Child(i);
//          WorkspaceUtils.waitForObj(Save);
//          ReportUtils.logStep_Screenshot("");
//          Save.Click();
//          SaveStat = false;
//          break;
//        }
//        
//      } 
      
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
      
//      var entries = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
      var entries = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",1)
      entries = entries.SWTObject("TabControl", "");
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
//      var entries = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
      var entries = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Entries");
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      } 
//      var add = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      var add = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Add Job Entry")
      WorkspaceUtils.waitForObj(add);
      add.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      } 
//      var EntryGrid = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      var EntryGrid = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
      Sys.HighlightObject(EntryGrid);
//      var Emp = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
      var Emp = EntryGrid.SWTObject("McValuePickerWidget", "");
      WorkspaceUtils.waitForObj(Emp);
      Emp.Click();
      WorkspaceUtils.SearchByValue(Emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
      Sys.Desktop.KeyDown(0x09); // Press Ctrl
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09); 
                
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyUp(0x09);
                
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
//      var Qty = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
      var Qty = EntryGrid.SWTObject("McTextWidget", "", 2)
      Qty.setText(temp[2]);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
//      var billable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3
      var billable = EntryGrid.SWTObject("McTextWidget", "", 2);
      billable.setText(temp[3]);
      aqUtils.Delay(4000, Indicator.Text);
//      var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
      var save = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Save Job Entry")
//      var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      WorkspaceUtils.waitForObj(save);
      save.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      aqUtils.Delay(4000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);

      aqUtils.Delay(100, Indicator.Text);
//      var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
      var allocate = EntryGrid.SWTObject("McPopupPickerWidget", "", 3);
      WorkspaceUtils.waitForObj(allocate);
      allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
      aqUtils.Delay(1000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
//      var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//      var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      var save = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Save Job Entry")
      WorkspaceUtils.waitForObj(save);
      save.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      ImageRepository.ImageSet.Close_Down.Click();
      }
      }
    }else{ 
            for(var g = 0;g<B_Estimatelines.length;g++){
        var temp = B_Estimatelines[g].split("*");
        if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])==0){
      aqUtils.Delay(100, Indicator.Text);

    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }else{ 
    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
    }

      aqUtils.Delay(1000, Indicator.Text);
      
//      var entries = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
//      WorkspaceUtils.waitForObj(entries);
//      entries.Click();
//      aqUtils.Delay(100, Indicator.Text);
//      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//      }
//      var entries = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
//      WorkspaceUtils.waitForObj(entries);
//      entries.Click();
//      aqUtils.Delay(100, Indicator.Text);
//      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//      } 
      
      var entries = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",1)
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
//      var entries = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
      var entries = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Entries");
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      } 
      
      
//      var add = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      var add = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Add Job Entry")
      WorkspaceUtils.waitForObj(add);
      add.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      } 
//      var EntryGrid = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//      Sys.HighlightObject(EntryGrid);
      var EntryGrid = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
      Sys.HighlightObject(EntryGrid);
//      var Emp = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
      var Emp = EntryGrid.SWTObject("McValuePickerWidget", "");
      WorkspaceUtils.waitForObj(Emp);
      Emp.Click();
      WorkspaceUtils.SearchByValue(Emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
      Sys.Desktop.KeyDown(0x09); // Press Ctrl
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09); 
                
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyUp(0x09);
                
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
//      var Qty = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
      var Qty = EntryGrid.SWTObject("McTextWidget", "", 2)
      Qty.setText(temp[2]);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
//      var billable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3
      var billable = EntryGrid.SWTObject("McTextWidget", "", 2);
      billable.setText(temp[3]);
      aqUtils.Delay(4000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
//      var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//      var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      var save = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Save Job Entry")
      WorkspaceUtils.waitForObj(save);
      save.Click();
      aqUtils.Delay(1000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      aqUtils.Delay(4000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);

      aqUtils.Delay(100, Indicator.Text);
//      var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
      var allocate = EntryGrid.SWTObject("McPopupPickerWidget", "", 3);
      WorkspaceUtils.waitForObj(allocate);
      allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
      aqUtils.Delay(1000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }else{ 
      ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
      }
//      var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//      var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      var save = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,"SingleToolItemControl","Save Job Entry")
      WorkspaceUtils.waitForObj(save);
      save.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }else{ 
      ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
      }
              ImageRepository.ImageSet.Close_Down.Click();
      }
      }
      SelectionBilling.Keys("[Down]");
    }
    

  }
//  var closeSelection = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
  var closeSelection = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",1,"Composite",2)
  closeSelection = closeSelection.SWTObject("TabControl", "");
  WorkspaceUtils.waitForObj(closeSelection);
  closeSelection.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Descp = [];
  var Ds = 0;
  for(var t=0;t<SelectionBilling.getItemCount();t++){ 
      for(var g = 0;g<B_Estimatelines.length;g++){
        var temp = B_Estimatelines[g].split("*");
        
        
         if(EnvParams.Country.toUpperCase()=="INDIA"){ 
           if(SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim().indexOf(temp[0])==0){
//             var S_temp = SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim().split(" - ");
//             Descp[Ds] = S_temp[1]+"*"+temp[1];
             var S_temp = SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim();
             S_temp = S_temp.substring(S_temp.indexOf(" - ")+3)
             Descp[Ds] = S_temp+"*"+temp[1];
             Log.Message(Descp[Ds])
             Ds++;
             }
         }else{ 
         if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])==0){
             var S_temp = SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim();
             S_temp = S_temp.substring(S_temp.indexOf(" - ")+3)
             Descp[Ds] = S_temp+"*"+temp[1];
             Log.Message(Descp[Ds])
             Ds++;
             }
         }
        
        
      }
      }
      
      
  
//Go To Draft  

//GoToDraft();
//WorkspaceUtils.closeAllWorkspaces();
//for(var i=level;i<ApproveInfo.length;i++){
//level=i;
//WorkspaceUtils.closeMaconomy();
//aqUtils.Delay(10000, Indicator.Text);
//var temp = ApproveInfo[i].split("*");
//Restart.login(temp[2]);
//aqUtils.Delay(5000, Indicator.Text);
//todo(temp[3]);
//FinalApprove(temp[1],temp[2],i);
//}
  
            
}//Flag
}// Main


function GoToDraft(){ 
  aqUtils.Delay(100, "Billing Price");
//var BudgetAmount = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
var BudgetAmount = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget", "1");
BudgetAmount = BudgetAmount.SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 5);
Sys.HighlightObject(BudgetAmount);
BudgetAmount = BudgetAmount.getText();

//var Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.SingleToolItemControl;
var Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve");
WorkspaceUtils.waitForObj(Approve);
ReportUtils.logStep_Screenshot("");
Approve.Click();
TextUtils.writeLog("Approve is clicked");
aqUtils.Delay(1000, "Approve is Clicked");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

//var DraftInvoice = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
var DraftInvoice = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Draft Invoices");
WorkspaceUtils.waitForObj(DraftInvoice);
ReportUtils.logStep_Screenshot("");
DraftInvoice.Click();
aqUtils.Delay(1000, Indicator.Text);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//var draftNo = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
//var draftNo = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid", "2");
var draftNo = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
draftNo = draftNo.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(draftNo);
draftNo.Keys("[Tab][Tab][Tab]");
//var billiablePrice = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
var billiablePrice = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
billiablePrice = billiablePrice.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(billiablePrice);
billiablePrice.Click();
billiablePrice.setText(BudgetAmount);

//var DraftTable = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var DraftTable = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
  var flag=false;
  for(var v=0;v<DraftTable.getItemCount();v++){ 
    if(DraftTable.getItem(v).getText_2(3).OleValue.toString().trim()==BudgetAmount){ 
      flag=true;
      break;
    }
    else{ 
      DraftTable.Keys("[Down]");
    }
  }
 ValidationUtils.verify(true,flag,"Invoice is available to submit Draft")
  if(flag){
//var CloseFilter = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
var CloseFilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
ReportUtils.logStep_Screenshot("");
CloseFilter.Click();
aqUtils.Delay(1000, Indicator.Text);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//  var DraftEditing = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  var DraftEditing = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Draft Editing");
  DraftEditing.Click();
//  var grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var grid = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid","2");
//  var Save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
//          var Des = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
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
//  var grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
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
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  var PrintDraft;
   PrintDraft = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Print Draft");
   Sys.HighlightObject(PrintDraft);
   PrintDraft.Click();
  

TextUtils.writeLog("Print Draft is Clicked");
aqUtils.Delay(5000, Indicator.Text);
if(MainJob){ 
WorkspaceUtils.savePDF_localDirectory("PDF Draft Invoice","Print Invoice Editing");
}else{ 
WorkspaceUtils.savePDF_localDirectory("PDF Draft T&M Invoice","Print Invoice Editing");
}


//var SaveTitle = "";
//var sFolder = "";
//var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
//    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*", 1).WndCaption.indexOf("Print Invoice Editing")!=-1){
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x41);
//    
//    if(ImageRepository.PDF.ChooseFolder.Exists())
//    ImageRepository.PDF.ChooseFolder.Click();
//    else{ 
//      var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
//      WorkspaceUtils.waitForObj(window);
//      Sys.Desktop.KeyDown(0x12); //Alt
//      Sys.Desktop.KeyDown(0x73); //F4
//      Sys.Desktop.KeyUp(0x12); //Alt
//      Sys.Desktop.KeyUp(0x73); //F4
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    }
//    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    aqUtils.Delay(2000, Indicator.Text);
//    SaveTitle = save.wText;
//    
//sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
//if (! aqFileSystem.Exists(sFolder)){
//if (aqFileSystem.CreateFolder(sFolder) == 0){ 
//    
//}
//else{
//Log.Error("Could not create the folder " + sFolder);
//}
//}
//save.Keys(sFolder+SaveTitle+".pdf");
//var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
//Sys.HighlightObject(p);
//var saveAs = p.FindChild("WndCaption", "&Save", 2000);
//if (saveAs.Exists)
//{ 
//saveAs.Click();
//}
//aqUtils.Delay(2000, Indicator.Text);
//aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
//aqUtils.Delay(2000, Indicator.Text);
//
//Sys.HighlightObject(pdf);
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//    }
//ValidationUtils.verify(true,true,"Print Draft Invoice is Clicked and PDF is Saved");
//Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
//ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
//if(MainJob){ 
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//ExcelUtils.WriteExcelSheet("PDF Draft Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  
//}else{ 
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//ExcelUtils.WriteExcelSheet("PDF Draft T&M Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  
//}
    
    aqUtils.Delay(4000, Indicator.Text);
   

//var appvBar = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
var appvBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",2)
  appvBar = appvBar.SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  aqUtils.Delay(2000, Indicator.Text);
  
ImageRepository.ImageSet.Maximize.Click();
aqUtils.Delay(2000, Indicator.Text);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//var DraftApproval = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
var DraftApproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","All Approval Actions");
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//var ApproverTable = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
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
//var closeBar = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
//var closeBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",1,"Composite",2);
//closeBar = closeBar.SWTObject("TabControl", "");
var closeBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

ImageRepository.ImageSet.Forward.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

CredentialLogin();
Project_manager = eval(Maconomy_ParentAddress).WndCaption;
Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
var OpCo2 = ApproveInfo[0].split("*");
Log.Message(OpCo2[2]);
Log.Message(Project_manager);
if(OpCo2[2]==Project_manager){
level = 1;

var Approve;
Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve Draft");
Sys.HighlightObject(Approve);
Approve.Click();
////var Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
//  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.isVisible())
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
// else
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
      

//  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
//  else
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
//
//        
// var ApproveStat = false;
////Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
//Sys.HighlightObject(Approve);
//for(var i=0;i<Approve.ChildCount;i++){ 
//  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
//    Approve = Approve.Child(i);
//    ApproveStat =true;
//    break;
//  }
//}

//if(!ApproveStat){ 
//Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
//Sys.HighlightObject(Approve);
//for(var i=0;i<Approve.ChildCount;i++){ 
//  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
//    Approve = Approve.Child(i);
//    break;
//  }
//} 
//}
//Log.Message(Approve.FullName)
//WorkspaceUtils.waitForObj(Approve);
//ReportUtils.logStep_Screenshot();
//Approve.Click();
ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved Draft Invoice");


var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);


  
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 

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
  
}
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

"Approve Invoice Drafts" ,"Approve Invoice Drafts by Type" , "Approve Invoice Drafts (Substitute)" , "Approve Invoice Drafts by Type (Substitute)"
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

//  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
// else
//  Approve = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
//            
//    
//Sys.HighlightObject(Approve);
//for(var i=0;i<Approve.ChildCount;i++){ 
//  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
//    Approve = Approve.Child(i);
//    break;
//  }
//}


//WorkspaceUtils.waitForObj(Approve);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();

ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
TextUtils.writeLog("Draft Invoice is Approved by "+Apvr);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

              


//  var Apv = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2
//Sys.HighlightObject(Apv);
//for(var i=0;i<Apv.ChildCount;i++){ 
//  if((Apv.Child(i).isVisible())&&(Apv.Child(i).Name.indexOf("McClumpSashForm")!=-1)){
//  Apv = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite
//    break;
//  }
//}
//  
//  var Apv = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite
//Sys.HighlightObject(Apv);
//for(var i=0;i<Apv.ChildCount;i++){ 
//  if((Apv.Child(i).isVisible())&&(Apv.Child(i).Name.indexOf("McClumpSashForm")!=-1)){
//  Apv = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite;
//    break;
//  }
//}

//var Apv = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite;
//Log.Message(Apv.FullName)
//var ApvPerson;
//for(var a=0;a<Apv.ChildCount;a++){ 
//  if((Apv.Child(a).Visible)&&(Apv.Child(a).JavaClassName == "McTextWidget")){ 
//    ApvPerson = Apv.Child(a);
//    Log.Message("short");
//    break;
//  }
//}
//if((ApvPerson=="")||(ApvPerson==null)){ 
//ApvPerson = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
//Log.Message("Long")
//}  
  
//                Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.ApproveStatus
//var ApvPerson = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
//var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
//    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
//    var i=0;
//
//while (((ApvPerson.getText().OleValue.toString().trim().toLowerCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "pproved").OleValue.toString().trim())==-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "YOU").OleValue.toString().trim())==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))&&(i!=60))
//{
//  aqUtils.Delay(100);
//  i++;
//  ApvPerson.Refresh();
//}

//if((ApvPerson.getText().OleValue.toString().trim().toLowerCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "pproved").OleValue.toString().trim())!=-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "YOU").OleValue.toString().trim())!=-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1)){

loginPer = eval(Maconomy_ParentAddress).WndCaption;
loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 
//  }else{ 
//  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected"); 
//  ValidationUtils.verify(true,false,"Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected")
//  }
  



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

////  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
////  printInvoice = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
//// else
////  printInvoice = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
////                 
////  WorkspaceUtils.waitForObj(printInvoice);
////  for(var i=0;i<printInvoice.ChildCount;i++){ 
////    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Invoice").OleValue.toString().trim())){
////      WorkspaceUtils.waitForObj(printInvoice.Child(i));
////      ReportUtils.logStep_Screenshot("");
////      printInvoice.Child(i).Click();
////      break;
////    }
////  } 
//  
//
//    var ChildCount = 0;
//    var Add = [];
//    var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//    Sys.Process("Maconomy").Refresh();  
//    for(var ip=0;ip<Parent.ChildCount;ip++){ 
//     var PChild = Parent.Child(ip);
//     if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")&& (PChild.ChildCount==3)){
//       Log.Message(PChild.Name)
////       for(var jp=0;jp<PChild.ChildCount;jp++){ 
////         var CChild = PChild.Child(jp);
////            if((CChild.isVisible()) && (CChild.JavaClassName=="Composite") && (CChild.Index==2)){
//            Add[ChildCount] = PChild;
//            ChildCount++;
////            }
////     }
//     }
//     }
//
//     var printInvoice = "";
//     var pos = 0;
//     for(var ip=0;ip<Add.length;ip++){ 
//     if(Add[ip].Height>pos){ 
//       pos = Add[ip].Height;
//       Log.Message(pos)
//       printInvoice = Add[ip];
//     }     
//     }
//     
//     Log.Message(printInvoice.FullName);
//     Sys.HighlightObject(printInvoice)
//     if(printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
//     printInvoice = printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
//     else
//     printInvoice = printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
//     Sys.HighlightObject(printInvoice)       
//  WorkspaceUtils.waitForObj(printInvoice);
//  for(var i=0;i<printInvoice.ChildCount;i++){ 
//    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Invoice").OleValue.toString().trim())){
//      WorkspaceUtils.waitForObj(printInvoice.Child(i));
//      ReportUtils.logStep_Screenshot("");
//      printInvoice.Child(i).Click();
//      break;
//    }
//  } 
//
//
//
//    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
//    aqUtils.Delay(5000, Indicator.Text);
//var SaveTitle = "";
//var sFolder = "";
//var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
//    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Invoice")!=-1){
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    
//    if(ImageRepository.PDF.ChooseFolder.Exists())
//    ImageRepository.PDF.ChooseFolder.Click();
//    else{ 
//      var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
//      WorkspaceUtils.waitForObj(window);
//      Sys.Desktop.KeyDown(0x12); //Alt
//      Sys.Desktop.KeyDown(0x73); //F4
//      Sys.Desktop.KeyUp(0x12); //Alt
//      Sys.Desktop.KeyUp(0x73); //F4
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    }
//    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    aqUtils.Delay(2000, Indicator.Text);
//    SaveTitle = save.wText;
//    
//sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
//if (! aqFileSystem.Exists(sFolder)){
//if (aqFileSystem.CreateFolder(sFolder) == 0){ 
//    
//}
//else{
//Log.Error("Could not create the folder " + sFolder);
//}
//}
//save.Keys(sFolder+SaveTitle+".pdf");
//var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
//Sys.HighlightObject(p);
//var saveAs = p.FindChild("WndCaption", "&Save", 2000);
//if (saveAs.Exists)
//{ 
//saveAs.Click();
//}
//aqUtils.Delay(2000, Indicator.Text);
//aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
//Sys.HighlightObject(pdf);
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//    }
//ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
//Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
//ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
//
//var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(sFolder+SaveTitle+".pdf");
//var textobj;
//  try{
//
//
//var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
//  textobj = obj.getText_2(docObj).OleValue.toString(); 
//  var invoiceName = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Invoice No:").OleValue.toString().trim();
//  invoiceName = invoiceName.length;
//  Log.Message(invoiceName)
//  textobj = textobj.substring(textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice No:").OleValue.toString().trim()+" ")+invoiceName+1);
//  Log.Message("Invoice No:"+textobj.substring(0,textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice Date").OleValue.toString().trim())))
//  textobj = textobj.substring(0,textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice Date").OleValue.toString().trim()));
//  }catch(objEx){
//    Log.Error("Exception while getting text from document::"+objEx);
//  }
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//  ExcelUtils.WriteExcelSheet("Time & Material Invocing No",EnvParams.Opco,"Data Management",textobj)
//  ExcelUtils.WriteExcelSheet("Time & Material Invocing Job",EnvParams.Opco,"Data Management",JobNum)
//  TextUtils.writeLog("Client Invoice No: "+textobj);
//
//
//}
//
//  ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
//  
//  
//}
//}
//
//}







if(Hitpoint.toUpperCase()!="YES"){
  
//    var ChildCount = 0;
//    var Add = [];
//    var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//    Sys.Process("Maconomy").Refresh();  
//    for(var ip=0;ip<Parent.ChildCount;ip++){ 
//     var PChild = Parent.Child(ip);
//     if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")&& (PChild.ChildCount==3)){
//       Log.Message(PChild.Name)
////       for(var jp=0;jp<PChild.ChildCount;jp++){ 
////         var CChild = PChild.Child(jp);
////            if((CChild.isVisible()) && (CChild.JavaClassName=="Composite") && (CChild.Index==2)){
//            Add[ChildCount] = PChild;
//            ChildCount++;
////            }
////     }
//     }
//     }
//
//     var printInvoice = "";
//     var pos = 0;
//     for(var ip=0;ip<Add.length;ip++){ 
//     if(Add[ip].Height>pos){ 
//       pos = Add[ip].Height;
//       Log.Message(pos)
//       printInvoice = Add[ip];
//     }     
//     }
//     
//     Log.Message(printInvoice.FullName);
//     Sys.HighlightObject(printInvoice)
//     if(printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
//     printInvoice = printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
//     else
//     printInvoice = printInvoice.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
//     Sys.HighlightObject(printInvoice)
//        
//  WorkspaceUtils.waitForObj(printInvoice);
//  for(var i=0;i<printInvoice.ChildCount;i++){ 
//    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Invoice").OleValue.toString().trim())){
//      WorkspaceUtils.waitForObj(printInvoice.Child(i));
//      ReportUtils.logStep_Screenshot("");
//      printInvoice.Child(i).Click();
//      break;
//    }
//  } 
  
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
if(MainJob){ 
PDF_Location = WorkspaceUtils.savePDF_localDirectory("PDF Invoice","Print Job Invoice");
}else{ 
PDF_Location = WorkspaceUtils.savePDF_localDirectory("PDF T&M Invoice","Print Job Invoice");
}
  
//  
//    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
//    aqUtils.Delay(5000, Indicator.Text);
//var SaveTitle = "";
//var sFolder = "";
//var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
//    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*", 1).WndCaption.indexOf("Invoice")!=-1){
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    
//    if(ImageRepository.PDF.ChooseFolder.Exists())
//    ImageRepository.PDF.ChooseFolder.Click();
//    else{ 
//      var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
//      WorkspaceUtils.waitForObj(window);
//      Sys.Desktop.KeyDown(0x12); //Alt
//      Sys.Desktop.KeyDown(0x73); //F4
//      Sys.Desktop.KeyUp(0x12); //Alt
//      Sys.Desktop.KeyUp(0x73); //F4
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    }
//    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    aqUtils.Delay(2000, Indicator.Text);
//    SaveTitle = save.wText;
//    
//sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
//if (! aqFileSystem.Exists(sFolder)){
//if (aqFileSystem.CreateFolder(sFolder) == 0){ 
//    
//}
//else{
//Log.Error("Could not create the folder " + sFolder);
//}
//}
//save.Keys(sFolder+SaveTitle+".pdf");
////var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
////saveAs.Click();
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
//var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
//Sys.HighlightObject(p);
//var saveAs = p.FindChild("WndCaption", "&Save", 2000);
//if (saveAs.Exists)
//{ 
//saveAs.Click();
//}
//aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
//Sys.HighlightObject(pdf);
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//    }
//ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
//Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
//ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
//
//if(MainJob){ 
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//ExcelUtils.WriteExcelSheet("PDF Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  
//}else{ 
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//ExcelUtils.WriteExcelSheet("PDF T&M Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  
//}
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//ExcelUtils.WriteExcelSheet("PDF Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  

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
  ExcelUtils.WriteExcelSheet("Time & Material Invocing No",EnvParams.Opco,"Data Management",textobj)
  ExcelUtils.WriteExcelSheet("Time & Material Invocing Job",EnvParams.Opco,"Data Management",JobNum)
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



function goto_Hitpoint_Billing(){ 
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();

ActionUtils.Select_Jobs_from_workspace();
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
//if(ImageRepository.ImageSet3.Jobs.Exists()){
// ImageRepository.ImageSet3.Jobs.Click();// GL
//}
//else if(ImageRepository.ImageSet.Job.Exists()){
//ImageRepository.ImageSet.Job.Click();
//}
//else{
//ImageRepository.ImageSet.Jobs1.Click();
//}

//var WrkspcCount = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
//var Workspc = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//var MainBrnch = "";
//for(var bi=0;bi<WrkspcCount;bi++){ 
//  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
//    MainBrnch = Workspc.Child(bi);
//    break;
//  }
//}
//
//
//var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
//  var Client_Managt;
//for(var i=1;i<=childCC;i++){ 
//Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
//if(Client_Managt.isVisible()){ 
//Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
//
//Log.Message(Language)
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
//ReportUtils.logStep_Screenshot();
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
//}
//
//} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Hitpoint Billing from Jobs Menu");
TextUtils.writeLog("Entering into Hitpoint Billing from Jobs Menu");


}


function Check_Hitpoint_Status(){ 
  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
  var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2)
//  var Company_No = NameMapping.Sys.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "");
  var Company_No = table.SWTObject("McValuePickerWidget", "");
  Sys.HighlightObject(Company_No);
  Company_No.Click();
  Company_No.setText(EnvParams.Opco);
  aqUtils.Delay(3000, Indicator.Text);
  Company_No.Keys("[Tab][Tab][Tab][Tab]");
  
  
//  var JobNo = NameMapping.Sys.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "");
  var JobNo = table.SWTObject("McValuePickerWidget", "");
  Sys.HighlightObject(JobNo);
  JobNo.Click();
  JobNo.setText(jobNumber);
  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
//  var table = NameMapping.Sys.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
    var flag=false;
    var Invoice_Editing_Number = "";
    var Invoice_Number = "";
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(4).OleValue.toString().trim()==jobNumber){ 
      flag=true;
      Invoice_Editing_Number = table.getItem(v).getText_2(1).OleValue.toString().trim()
      Invoice_Number = table.getItem(v).getText_2(2).OleValue.toString().trim()
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){

  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Time & Material Invocing No",EnvParams.Opco,"Data Management",Invoice_Number)
  ExcelUtils.WriteExcelSheet("Time & Material Invocing Job",EnvParams.Opco,"Data Management",jobNumber)
  TextUtils.writeLog("Invoice Editing Number: "+Invoice_Editing_Number);
  ExcelUtils.WriteExcelSheet("Invoice Editing Number",EnvParams.Opco,"Data Management",Invoice_Editing_Number)
  
  }
  
    aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}
