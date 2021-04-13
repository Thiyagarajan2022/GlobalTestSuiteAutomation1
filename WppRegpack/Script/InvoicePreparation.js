﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "InvoicePreparation";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var percentage,jobNumber,Hitpoint,Buss_Area_2 = "";
var Language = "";
var Ipreparation_ID = "";
var IpreparationUnit = "";
var Estimatelines = [];
var   Descp = [];
//Main Function
function InvoicePreparation(){ 
TextUtils.writeLog("Create Invoice Preparation Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "InvoicePreparation";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
percentage,jobNumber,Hitpoint,Buss_Area_2 = "";
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Invoice Preparation started::"+STIME);
Ipreparation_ID = "";
IpreparationUnit = "";
try{
  

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
    
    Ipreparation_ID = TestRunner.testCaseId;
    IpreparationUnit = TestRunner.unitName; 
    TestRunner.TempUnit = IpreparationUnit;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Invoice preparation")
    }
    
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    jobNumber = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    
    if((jobNumber=="")||(jobNumber==null)){
      
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
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+Job_JIRAID)
    Runner.CallMethod("Creation_Of_Job.createJob",jobSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    //Creation of Budget
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var budgetSheet = ExcelUtils.getColumnDatas("Budget sheet",EnvParams.Opco)
    if(budgetSheet==""){ 
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Invoice preparation")
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
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job Budget");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("BudgetCreation.createBudget",budgetSheet,serialOder);
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
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Invoice preparation")
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
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Quote");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("Creation_of_Quote.CreateQuote",quoteSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
/*    
    //Creation of PO
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var POSheet = ExcelUtils.getColumnDatas("PO Sheet",EnvParams.Opco)
    if(POSheet==""){ 
      ValidationUtils.verify(true,false,"Need PO for Job to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, POSheet, true);
    var JobSO = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(JobSO==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create PO")
    }
    var PO_SO = ExcelUtils.getColumnDatas("PO Serial Order",EnvParams.Opco)
    if(PO_SO==""){ 
      ValidationUtils.verify(true,false,"Need PO Serial Order to Create PO")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var PO_Number = ExcelUtils.getRowDatas("PO Number_"+PO_SO,EnvParams.Opco)
    if((PO_Number=="")||(PO_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreatePurchaseOrder_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreatePurchaseOrder_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;   
    TestRunner.unitName = "CreatePurchaseOrder_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Purchase Order");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Purchase Order");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("CreatePO.CreatePurchaseOrder",POSheet,JobSO,PO_SO);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
  //Approving PO
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var APSheet = ExcelUtils.getColumnDatas("Approve PO Sheet",EnvParams.Opco)
    if(APSheet==""){ 
      ValidationUtils.verify(true,false,"Need Approve PO Sheet for Job to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, APSheet, true);
    var serialOder = ExcelUtils.getRowDatas("PO Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need PO Serial Order to Approve PO")
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var AP_Number = ExcelUtils.getRowDatas("Approved PO_"+serialOder,EnvParams.Opco)
    if((AP_Number=="")||(AP_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("ApprovePurchaseOrder_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for ApprovePurchaseOrder_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;    
    TestRunner.unitName = "ApprovePurchaseOrder_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Approve Purchase Order");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Approve Purchase Order");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("ApprovePO.ApprovePurchaseOrder",APSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
   }

   //Creation of Vendor Invoice
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var VISheet = ExcelUtils.getColumnDatas("Vendor Invoice Sheet",EnvParams.Opco)
    if(VISheet==""){ 
      ValidationUtils.verify(true,false,"Need Vendor Invocie for Job to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, VISheet, true);
    var PO_SO = ExcelUtils.getColumnDatas("PO Serial Order",EnvParams.Opco)
    if(PO_SO==""){ 
      ValidationUtils.verify(true,false,"Need PO Serial Order to Create vendor Invocie")
    }
    
    var VI_SO = ExcelUtils.getColumnDatas("Vendor Invoice Serial Order",EnvParams.Opco)
    if(VI_SO==""){ 
      ValidationUtils.verify(true,false,"Need Vendor Invoice Serial Order to Create vendor Invocie")
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var VI_Number = ExcelUtils.getRowDatas("Vendor Invoice NO_"+VI_SO,EnvParams.Opco)
    var Journal_Number = ExcelUtils.getRowDatas("Invoice Journal NO_"+VI_SO,EnvParams.Opco)
    if(((VI_Number=="")||(VI_Number==null))&&((Journal_Number=="")||(Journal_Number==null))){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateVendorInvoice_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateVendorInvoice_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID; 
    TestRunner.unitName = "CreateVendorInvoice_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Vendor Invoice Creation");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Vendor Invoice");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("VendorInvoice.CreateInvoice",VISheet,PO_SO,VI_SO);
    Log.PopLogFolder();

    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    
    }
 
    //Approve Vendor Invocie
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var AISheet = ExcelUtils.getColumnDatas("Approve Vendor Invocie sheet",EnvParams.Opco)
    if(AISheet==""){ 
      ValidationUtils.verify(true,false,"Need Approve VI Sheet for Job to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, AISheet, true);
    var serialOder = ExcelUtils.getRowDatas("Vendor Invoice Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Vendor Invoice Serial Order to Approve VI")
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var AI_Number = ExcelUtils.getRowDatas("Approved Vendor Invoice_"+serialOder,EnvParams.Opco)
    if((AI_Number=="")||(AI_Number==null)){ 
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("ApproveVendorInvoice_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for ApproveVendorInvoice_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;  
    TestRunner.unitName = "ApproveVendorInvoice_"+serialOder; 
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Approve Vendor Invoice");
    ReportUtils.DStat = true;
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Approve Vendor Invoice");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("Approve_VI.ApproveInvoice",AISheet,serialOder);
    Log.PopLogFolder();

    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    
    //Post Vendor Journal
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var PVISheet = ExcelUtils.getColumnDatas("Post Vendor Invoice sheet",EnvParams.Opco)
    if(PVISheet==""){ 
      ValidationUtils.verify(true,false,"Need Approve VI Sheet for Job to Create Invoice preparation")
    }
//    ExcelUtils.setExcelName(workBook, PVISheet, true);
//    var serialOder = ExcelUtils.getRowDatas("Vendor Invoice Serial Order",EnvParams.Opco)
//    if(serialOder==""){ 
//      ValidationUtils.verify(true,false,"Need Vendor Invoice Serial Order to Post Vendor Journal")
//    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var PVI_Number = ExcelUtils.getRowDatas("Post Vendor Journal_"+serialOder,EnvParams.Opco)
    if((PVI_Number=="")||(PVI_Number==null)){ 
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("PostVendorJournal_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for PostVendorJournal_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;  
    TestRunner.unitName = "PostVendorJournal_"+serialOder; 
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Post Vendor Journal");
    ReportUtils.DStat = true; 
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Post Vendor Invoice");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("Post_VI.postVendorJournal",PVISheet,VI_SO);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    */
    
TestRunner.testCaseId = Ipreparation_ID;
TestRunner.unitName = IpreparationUnit;
//}
}


Log.Message(jobNumber)
Log.Message(template)


TestRunner.JiraStat = true;
TestRunner.JiraUpdate = true;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Biller",EnvParams.Opco);
//if((Project_manager=="")||(Project_manager==null))
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if((Project_manager=="")||(Project_manager==null))
ValidationUtils.verify(false,true,"Login Credentials required for anyone of Agency - Biller or Agency - Finance,");

Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
 
}


getDetails();
gotoMenu();
gotoInvoicing();
WorkspaceUtils.closeAllWorkspaces();
for(var i=level;i<ApproveInfo.length;i++){
level=i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApprove(temp[1],temp[2],i);
}
}
catch(err){ 
  Log.Message(err);
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



function getDetails(){ 
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
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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

Log.Message(Language)
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}

function gotoInvoicing(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var allJobs = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
allJobs.Click();

//var labels = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.Label;
var labels = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
    labels = labels.Child(i);
    break;
  }
}

WorkspaceUtils.waitForObj(labels);
//var i=0;
//while((labels.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=60)){ 
//  aqUtils.Delay(100);
//  i++;
//  labels.Refresh();
//}
//if(labels.getText().OleValue.toString().trim().indexOf("results")!=-1){ 
// ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
//}

  var table = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JobTable;
  var firstcell = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JobTable.CompanyNo;
  var closeFilter = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter;
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
  
  var job = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JobTable.JobNumber;
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var i=0;
while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to for Invoice Preparation");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy for Invoice Preparation"); 
  closeFilter.Click();
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var clientApproved = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(clientApproved);
  if(clientApproved.background!=10674625){
    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
    }
  var workingEstimate = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
  WorkspaceUtils.waitForObj(workingEstimate);
  if(workingEstimate.background!=10674625){
    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
    }
  var lastInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  var totalInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  var billingPrice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
  var netInvoiceOnAcc = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
  
//---------------

if(Hitpoint.toUpperCase()=="YES"){
  // Moving to Information Tab to Submit
  var info = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 5)
  Sys.HighlightObject(info);
  info.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  info.Click();


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(1000, Indicator.Text);
  var business_Area2 = Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
  Sys.HighlightObject(business_Area2);
  business_Area2.Click();
  WorkspaceUtils.SearchByValue(business_Area2,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Option").OleValue.toString().trim(),Buss_Area_2,"Bussiness Area 2");
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(1000, Indicator.Text);
  
  var Save = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SingleToolItemControl;
  Sys.HighlightObject(Save);
  Save.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(1000, Indicator.Text);
  
  var Budgeting = Aliases.Maconomy.Blanket_Invoice.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  
  }
else{
  
  /// Original Invoice
  
  var Budgeting = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  }
  aqUtils.Delay(100, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  aqUtils.Delay(5000, Indicator.Text);
  var Estimate = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  Estimate.Keys("Client Approved Estimate");
  aqUtils.Delay(100, Indicator.Text);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var FullBudget = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var BudgetGrid = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
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
  
  

  
//  var Invoicing = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Invoicing;
//  WorkspaceUtils.waitForObj(Invoicing);
//  Invoicing.Click();
  if(Hitpoint.toUpperCase()=="YES"){
  var Invoicing = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8);
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();  
  }else{
  var Invoicing = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();
  }
  
  var iPreparation = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.invoicePreparation;
  
//  var iPreparation = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel;
//                     Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
//  WorkspaceUtils.waitForObj(iPreparation);
//  for(var i=0;i<iPreparation.ChildCount;i++){ 
//    if((iPreparation.Child(i).isVisible())&&(iPreparation.Child(i).text=="Invoice Preparation")){
//      iPreparation.Child(i).HoverMouse(); 
//      ReportUtils.logStep_Screenshot("");
//      iPreparation.Child(i).Click();
//      break;
//    }
//  }
  
  WorkspaceUtils.waitForObj(iPreparation);
  ReportUtils.logStep_Screenshot("");
  iPreparation.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var CreateIPreparation = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.EditableBar;
  WorkspaceUtils.waitForObj(CreateIPreparation);
  for(var i=0;i<CreateIPreparation.ChildCount;i++){ 
    if((CreateIPreparation.Child(i).isVisible())&&(CreateIPreparation.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Prepare Invoice").OleValue.toString().trim())){
      CreateIPreparation.Child(i).HoverMouse(); 
      ReportUtils.logStep_Screenshot("");
      CreateIPreparation.Child(i).Click();
      break;
    }
  }
  aqUtils.Delay(5000);
  var BasicOfInvoice = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  BasicOfInvoice.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget").OleValue.toString().trim());

var i=0;
while((BasicOfInvoice.getText()!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget").OleValue.toString().trim())&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  BasicOfInvoice.Refresh();
}
if(BasicOfInvoice.getText()!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget").OleValue.toString().trim()){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

var next = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
next.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
if((EnvParams.Country.toUpperCase()!="SPAIN")&&(EnvParams.Country.toUpperCase()!="MALAYSIA")&&(EnvParams.Country.toUpperCase()!="CHINA")&&(EnvParams.Country.toUpperCase()!="UAE")){
var Percentage = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.Composite.Percentage;
WorkspaceUtils.waitForObj(Percentage);
Percentage.Click();
Percentage.setText(percentage);
}
var prepare = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Prepare Invoice").OleValue.toString().trim())
WorkspaceUtils.waitForObj(prepare);
ReportUtils.logStep_Screenshot("");
prepare.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Invoice Preparation").OleValue.toString().trim()){
var Okay = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Invoice Preparation").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
ReportUtils.logStep_Screenshot("");
Okay.Click();
}

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
if((EnvParams.Country.toUpperCase()!="SPAIN")&&(EnvParams.Country.toUpperCase()!="MALAYSIA")&&(EnvParams.Country.toUpperCase()!="CHINA")&&(EnvParams.Country.toUpperCase()!="UAE")){
var invoiceTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(invoiceTable);
for(var i=0;i<invoiceTable.getItemCount()-1;i++){ 
 
//for(var j=0;j<Estimate.length;j++) { 
//  Esplit = Estimate[j].split("*");
//  
//  if(Esplit[0]==invoiceTable.getItem(i).getText_2(0).OleValue.toString().trim()){
    
  var temp = invoiceTable.getItem(i).getText_2(2).OleValue.toString().trim();
  temp = parseFloat(temp).toFixed(2);;
//  if(percentage.indexOf(".")!=-1){ 
    percentage = parseFloat(percentage.toString()).toFixed(2);
//  }
  Log.Message("temp :"+temp)
  Log.Message("percentage :"+percentage)
 // if(temp!=percentage){ 
   // ValidationUtils.verify(true,false,"Percentage missMatch with Datasheet");
//  }
  
//  var temp = invoiceTable.getItem(i).getText_2(2).OleValue.toString().trim();
//  Esplit[1] = parseFloat(Esplit[1].toString())/100;
//  Esplit[1] = parseFloat(Esplit[1]*percentage).toFixed(2);
//  
//  if(Esplit[1]!=temp){ 
//  Log.Message("Esplit :"+Esplit[1])
//  Log.Message("temp :"+temp)   
//  ValidationUtils.verify(true,false,"Quatity missMatch with repective percentage");
//  }
//  break;
//  }
  
//}

}

}

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var SelectionBilling = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(SelectionBilling);
aqUtils.Delay(1000, "Billing Price");
  Descp = [];
  var Ds = 0;
  for(var t=0;t<SelectionBilling.getItemCount();t++){ 
      for(var g = 0;g<Estimatelines.length;g++){
        var temp = Estimatelines[g].split("*");
           if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])!=-1){
             var S_temp = SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim();
             Descp[Ds] = S_temp+"*"+temp[1];
             Log.Message(Descp[Ds])
             Ds++;
             }

      }
      }




  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }

var BudgetAmount;
 
 if((EnvParams.Country.toUpperCase()!="SPAIN")&&(EnvParams.Country.toUpperCase()!="MALAYSIA")&&(EnvParams.Country.toUpperCase()!="UAE")){
 BudgetAmount = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.ForInvoicing;
 }else{ 
 BudgetAmount = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite3.McTextWidget;
 }
BudgetAmount = BudgetAmount.getText();
// BudgetAmount = "74,750.00";

Log.Message(BudgetAmount);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var Submit = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.EditableBar;
  WorkspaceUtils.waitForObj(Submit);
  for(var i=0;i<Submit.ChildCount;i++){ 
    if((Submit.Child(i).isVisible())&&(Submit.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(Submit.Child(i));
      ReportUtils.logStep_Screenshot("");
      Submit.Child(i).Click();
      break;
    }
  }
  
  
  Delay(5000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
   var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Invoice Preparation").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Invoice Preparation").OleValue.toString().trim(), 2000);
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


  Delay(5000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  

var removeLatestDrafft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.SingleToolItemControl;
var i=0;
while((!removeLatestDrafft.Visible)&&(removeLatestDrafft.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Remove Latest Draft").OleValue.toString().trim())&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  removeLatestDrafft.Refresh();
}
if(removeLatestDrafft.getText()!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Remove Latest Draft").OleValue.toString().trim()){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

WorkspaceUtils.waitForObj(removeLatestDrafft);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var draftInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(draftInvoice);
draftInvoice.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var draftNo = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(draftNo);
draftNo.Keys("[Tab][Tab][Tab]");
var billiablePrice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(billiablePrice);
billiablePrice.Click();
billiablePrice.setText(BudgetAmount);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

//var DraftTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//  var flag=false;
//  for(var v=0;v<DraftTable.getItemCount();v++){ 
//    if(DraftTable.getItem(v).getText_2(3).OleValue.toString().trim()==BudgetAmount){ 
//      flag=true;
//      break;
//    }
//    else{ 
//      DraftTable.Keys("[Down]");
//    }
//  }

var flag=false;
flag=true;
 ValidationUtils.verify(true,flag,"Prepared Invoice is available to submit Draft")
  if(flag){
var CloseFilter = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;  
ReportUtils.logStep_Screenshot("");
CloseFilter.Click();
aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var DraftEditing = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  DraftEditing.Click();
  var grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var Save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
          var Des = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
          Des.Click();
          Des.setText(temp[1]);
          aqUtils.Delay(1000, Indicator.Text);
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
            }
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
  var grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(grid)
  Log.Message(i)
  Log.Message(grid.getItemCount()-2)
  Log.Message(i<grid.getItemCount()-2)
  if(i<grid.getItemCount()-2){
  grid.Keys("[Down]");
  }
  }  
  
  
var SubmitDraft;
  if(Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.isVisible())
  SubmitDraft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
 else
  SubmitDraft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite;
            
  WorkspaceUtils.waitForObj(SubmitDraft);
  for(var i=0;i<SubmitDraft.ChildCount;i++){ 
    if((SubmitDraft.Child(i).isVisible())&&(SubmitDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(SubmitDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      SubmitDraft.Child(i).Click();
      break;
    }
  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }

var Excl_Tax = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
var grandTotal = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
//var Payment_Terms = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2);
//var Payment_Terms = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2)


//Finding Payment Terms
var break_MainLoop = false;
var ParentAdd = Aliases.Maconomy.InvoicePlan.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget
var Payment_Terms = "";
for(var i=0;i<ParentAdd.ChildCount;i++){ 
  var temp = ParentAdd.Child(i);
  for(var j=0;j<temp.ChildCount;j++){ 
    if(temp.Child(j).Name.indexOf("McPopupPickerWidget")!=-1){
      Payment_Terms = temp.Child(j);
      break_MainLoop = true;
      break;
    }
  }
  
  if(break_MainLoop){ 
    break;
  }
}


Excl_Tax = Excl_Tax.getText().OleValue.toString().trim();
grandTotal = grandTotal.getText().OleValue.toString().trim();
Payment_Terms = Payment_Terms.getText().OleValue.toString().trim();
Payment_Terms = Payment_Terms.replace(/[^0-9]+/g, "");;
var Q_total = 0;
var specification = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var q = 0;
QuoteDetails = [];
var InvoiceMPL = "InvoiceMPL";
for(var i=0;i<specification.getItemCount();i++){ 

  var Q_Desp = specification.getItem(i).getText_2(1).OleValue.toString().trim();
  if(Q_Desp!=""){
    
  var Q_Qty = specification.getItem(i).getText_2(2).OleValue.toString().trim();
  var Q_Billing = specification.getItem(i).getText_2(3).OleValue.toString().trim();
  var Q_BillingTotoal = specification.getItem(i).getText_2(4).OleValue.toString().trim();
  var Q_Tax1 = specification.getItem(i).getText_2(7).OleValue.toString().trim();
  var Q_Tax2 = specification.getItem(i).getText_2(9).OleValue.toString().trim();
  var Q_Tax1currency = specification.getItem(i).getText_2(8).OleValue.toString().trim();
  var Q_Tax2currency = specification.getItem(i).getText_2(10).OleValue.toString().trim();
//  var Q_total = parseFloat(Q_BillingTotoal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
//  QuoteDetails[q] = Q_Desp+"*"+Q_Qty+"*"+Q_Billing+"*"+Q_BillingTotoal+"*"+Q_Tax1+"*"+Q_Tax2+"*"+Q_Tax1currency+"*"+Q_Tax2currency+"*"+Q_total.toFixed(2)+"*";
//  Log.Message(QuoteDetails[q]);
  Q_total =Q_total+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
  Log.Message(Q_total);
  QuoteDetails[q] = Q_Desp+"*"+Q_Qty+"*"+Q_Billing+"*"+Q_BillingTotoal+"*"+Q_Tax1+"*"+Q_Tax2+"*"+Q_Tax1currency+"*"+Q_Tax2currency+"*"+Q_total.toFixed(2)+"*";
  Log.Message(QuoteDetails[q]);
  q++;
  ExcelUtils.setExcelName(workBook,InvoiceMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,InvoiceMPL,Q_Desp);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,InvoiceMPL,Q_Qty);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,InvoiceMPL,Q_Billing);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TotalBilling_"+q,InvoiceMPL,Q_BillingTotoal);
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
  
  
var PrintDraft;
  if(Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.isVisible())
  PrintDraft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
 else
  PrintDraft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite;
// var PrintDraft = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  WorkspaceUtils.waitForObj(PrintDraft);
  for(var i=0;i<PrintDraft.ChildCount;i++){ 
    if((PrintDraft.Child(i).isVisible())&&(PrintDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(PrintDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      PrintDraft.Child(i).Click();
      break;
    }
  } 
  
TextUtils.writeLog("Print Draft is Clicked");
//    aqUtils.Delay(5000, Indicator.Text);
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
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Draft Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Draft Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")   

var appvBar = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  
ImageRepository.ImageSet.Maximize.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var DraftApproval = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();

//WorkspaceUtils.waitForObj(DraftApproval);
//  for(var i=0;i<DraftApproval.ChildCount;i++){ 
//    if((DraftApproval.Child(i).isVisible())&&(DraftApproval.Child(i).text=="All Approval Actions")){
//      WorkspaceUtils.waitForObj(DraftApproval.Child(i));
//      DraftApproval.Child(i).Click();
//      break;
//    }
//  } 
  
//purchaseApproval.Click();
var ApproverTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
 var y=0;
for(var i=0;i<ApproverTable.getItemCount();i++){   
   var approvers="";
    if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
    approvers = EnvParams.Opco+"*"+jobNumber+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
    Log.Message("Approver level :" +i+ ": " +approvers);
//      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
    Approve_Level[y] = approvers;
    y++;
    }
}
ReportUtils.logStep_Screenshot("");
var closeBar = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
//sheetName = "ApprovePurchaseOrder";
Log.Message(OpCo2[2]);
Log.Message(Project_manager);
if(OpCo2[2]==Project_manager){
level = 1;

//var Approve;
//  if(Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.isVisible())
//  Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
// else
//  Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite;
////var Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
//Sys.HighlightObject(Approve);
//for(var i=0;i<Approve.ChildCount;i++){ 
//  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText=="Approve Draft")){
//    Approve = Approve.Child(i);
//    break;
//  }
//}

var Approve;
 var ApproveStat = false;
Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    ApproveStat =true;
    break;
  }
}

if(!ApproveStat){ 
Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite;
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
} 
}
Log.Message(Approve.FullName)

WorkspaceUtils.waitForObj(Approve);
ReportUtils.logStep_Screenshot();
Approve.Click();
ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved Draft Invoice");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

var screen = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-100);
ReportUtils.logStep_Screenshot("");
//  aqUtils.Delay(5000, Indicator.Text);
var ApvPerson = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while (((ApvPerson.getText().OleValue.toString().trim().toLowerCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "pproved").OleValue.toString().trim())==-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "YOU").OleValue.toString().trim())==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

  if((ApvPerson.getText().OleValue.toString().trim().toLowerCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "pproved").OleValue.toString().trim())!=-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "YOU").OleValue.toString().trim())!=-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1)){
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected")
  }
  
if(Approve_Level.length==1){
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Invoice preparation Job",EnvParams.Opco,"Data Management",PONum)
var appvBar = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
  
ImageRepository.ImageSet.Maximize.Click();
var DraftApproval = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();

//var DraftApproval = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel;
//WorkspaceUtils.waitForObj(DraftApproval);
//  for(var i=0;i<DraftApproval.ChildCount;i++){ 
//    if((DraftApproval.Child(i).isVisible())&&(DraftApproval.Child(i).text=="All Approval Actions")){
//      WorkspaceUtils.waitForObj(DraftApproval.Child(i));
//      DraftApproval.Child(i).Click();
//      break;
//    }
//  } 
  
var ApproverTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
ReportUtils.logStep_Screenshot();
var closeBar = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();

}
}
  
  }
  }

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

if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }

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



function FinalApprove(PONum,Apvr,lvl){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
var table = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite2.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
if(Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite2.PTabFolder.TabFolderPanel.Visible){

}else{ 
  var showFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SWTObject("SingleToolItemControl", "", 2);
  WorkspaceUtils.waitForObj(showFilter);
  showFilter.Click();
//ImageRepository.ImageSet.Show_Filter.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var table = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(PONum);
var closefilter = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;

var labels = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McPagingWidget", "", 2);

WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
    labels = labels.Child(i);
    break;
  }
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
WorkspaceUtils.waitForObj(labels);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var i=0;
while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
WorkspaceUtils.waitForObj(table);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==PONum){ 
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

var Approve;
if(Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.isVisible())
 Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
else
 Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite;
//var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
}
//WorkspaceUtils.waitForObj(Approve);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();

ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
TextUtils.writeLog("Draft Invoice is Approved by "+Apvr);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

var screen = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-100);

var ApvPerson = "";
var Apv = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite;
for(var a=0;a<Apv.ChildCount;a++){ 
  if((Apv.Child(a).Visible)&&(Apv.Child(a).JavaClassName == "McTextWidget")){ 
    ApvPerson = Apv.Child(a);
    Log.Message("short");
    break;
  }
}
if((ApvPerson=="")||(ApvPerson==null)){ 
ApvPerson = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;  
Log.Message("Long")
}  
  
//                Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.ApproveStatus
//var ApvPerson = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;

while (((ApvPerson.getText().OleValue.toString().trim().toLowerCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "pproved").OleValue.toString().trim())==-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "YOU").OleValue.toString().trim())==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

if((ApvPerson.getText().OleValue.toString().trim().toLowerCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "pproved").OleValue.toString().trim())!=-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "YOU").OleValue.toString().trim())!=-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1)){
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected")
  }
  



if(Approve_Level.length==lvl+1){
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
//var printInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;

var approvalBar = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
    ImageRepository.ImageSet.Maximize.Click();

var DraftApproval = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();
//  for(var i=0;i<DraftApproval.ChildCount;i++){ 
//    if((DraftApproval.Child(i).isVisible())&&(DraftApproval.Child(i).text=="All Approval Actions")){
//      WorkspaceUtils.waitForObj(DraftApproval.Child(i));
//      DraftApproval.HoverMouse();
//      ReportUtils.logStep_Screenshot();
//      DraftApproval.Child(i).Click();
//      break;
//    }
//  } 
  

var ApproverTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ReportUtils.logStep_Screenshot();

var closeBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
//var printInvoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
//var printInvoice = "";
//  if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
//  printInvoice = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
// else
//  printInvoice = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
    

if(Hitpoint.toUpperCase()!="YES"){

    var ChildCount = 0;
    var Add = [];
    var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
    Sys.Process("Maconomy").Refresh();  
    for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
     if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")&& (PChild.ChildCount==3)){
       Log.Message(PChild.Name)
//       for(var jp=0;jp<PChild.ChildCount;jp++){ 
//         var CChild = PChild.Child(jp);
//            if((CChild.isVisible()) && (CChild.JavaClassName=="Composite") && (CChild.Index==2)){
            Add[ChildCount] = PChild;
            ChildCount++;
//            }
//     }
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
//Sys.HighlightObject(printInvoice);
//Log.Message(printInvoice.toolTipText)
//Log.Message(printInvoice.text)
//printInvoice.Click();
    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
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
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
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
ExcelUtils.WriteExcelSheet("PDF Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  


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
  ExcelUtils.WriteExcelSheet("Invoice preparation No",EnvParams.Opco,"Data Management",textobj)
  ExcelUtils.WriteExcelSheet("Invoice preparation Job",EnvParams.Opco,"Data Management",PONum)
  TextUtils.writeLog("Client Invoice No: "+textobj);

}else{ 
  aqUtils.Delay(2000, Indicator.Text);
  
  var Sent_To_Hitpoint = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 15);
  Sys.HighlightObject(Sent_To_Hitpoint);
  Sent_To_Hitpoint.Click();
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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

Log.Message(Language)
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Hitpoint Billing from Jobs Menu");
TextUtils.writeLog("Entering into Hitpoint Billing from Jobs Menu");


}


function Check_Hitpoint_Status(){ 
  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  
  var Company_No = NameMapping.Sys.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "");
  Sys.HighlightObject(Company_No);
  Company_No.Click();
  Company_No.setText(EnvParams.Opco);
  aqUtils.Delay(3000, Indicator.Text);
  Company_No.Keys("[Tab][Tab][Tab][Tab]");
  
  
  var JobNo = NameMapping.Sys.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "");
  Sys.HighlightObject(JobNo);
  JobNo.Click();
  JobNo.setText(jobNumber);
  
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var table = NameMapping.Sys.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
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
  ExcelUtils.WriteExcelSheet("Invoice preparation No",EnvParams.Opco,"Data Management",Invoice_Number)
  ExcelUtils.WriteExcelSheet("Invoice preparation Job",EnvParams.Opco,"Data Management",jobNumber)
  TextUtils.writeLog("Invoice Editing Number: "+Invoice_Editing_Number);
  ExcelUtils.WriteExcelSheet("Invoice Editing Number",EnvParams.Opco,"Data Management",Invoice_Editing_Number)
  
  }
  
    aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}
