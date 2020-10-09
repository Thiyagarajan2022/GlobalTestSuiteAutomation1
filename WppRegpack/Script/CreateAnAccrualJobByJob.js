//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateAnAccrualJobByJob";
var Language = "";
Indicator.Show();
  
/** 
 * This script Create Accural for Single PO
 * @author  : Muthu Kumar M
 * @version : 1.0
 * Created Date :10/07/2020
 */
 
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var JobNo,WorkCode,EntryDate,NoForAccrual,PoNumber ;

// Variables used to assign TestCaseID for dependency Job in JIRA
var Ipreparation_ID = "";
var IpreparationUnit = "";

/**
  *  This Main function invokes maconomy and calls subfunctionality methods
  */
function CreateAnAccrualJobByJob() {
TextUtils.writeLog("Create An Accrual Job By Job Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateAnAccrualJobByJob";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
JobNo,WorkCode,EntryDate,NoForAccrual,PoNumber ="";


STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);
aqUtils.Delay(3000, Indicator.Text);

//Getting Details from Datasheet
getDetails();

//Checking Login to execute Accural JOb by Job
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  
}
goToJobMenuItem();   
GoToAccruals();

//Close all Open Workspace in Maconomy
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
aqTestCase.End();

}

//getting data from datasheet
function getDetails(){
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  JobNo = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  var invoicePreparation = ExcelUtils.getRowDatas("Invoice preparation Job",EnvParams.Opco);
  var AllocationWIP = ExcelUtils.getRowDatas("Job Invoice Allocation with WIP Job",EnvParams.Opco);
  var invoiceBudget = ExcelUtils.getRowDatas("Invoice from Budget Job",EnvParams.Opco);
  var invoiceAccount = ExcelUtils.getRowDatas("Invoice OnAccount Job",EnvParams.Opco);
  var writeoffInvoice = ExcelUtils.getRowDatas("Write Off Invoicing Job",EnvParams.Opco);
  
// Checking Job Number has been used any of Invoice
  if(((JobNo=="")||(JobNo==null))||(invoicePreparation==JobNo)||(AllocationWIP==JobNo)||(invoiceBudget==JobNo)||(invoiceAccount==JobNo)||(writeoffInvoice==JobNo)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  JobNo = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
// Making Job Number as ' Null ' to Create New Job if existing Job Number has used for any of Invoice
  if((invoicePreparation==JobNo)||(AllocationWIP==JobNo)||(invoiceBudget==JobNo)||(invoiceAccount==JobNo)||(writeoffInvoice==JobNo)){
   JobNo = "";
   Create_New_Job();
  }else{ 
    
ExcelUtils.setExcelName(workBook, "Data Management", true);
PoNumber = ReadExcelSheet("PO Number",EnvParams.Opco,"Data Management");
if((PoNumber=="")||(PoNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  PoNumber = ExcelUtils.getRowDatas("PoNumber",EnvParams.Opco)
  }
  if((PoNumber==null)||(PoNumber=="")){ 
  ValidationUtils.verify(false,true,"PoNo is Needed to Create a Job");
  }
  Log.Message(PoNumber)  
  
  }
  
Log.Message(JobNo)
ExcelUtils.setExcelName(workBook, sheetName, true);
WorkCode = ExcelUtils.getColumnDatas("WorkCode",EnvParams.Opco)
if((WorkCode==null)||(WorkCode=="")){ 
ValidationUtils.verify(false,true,"WorkCode is Needed to Create a Job");
}
Log.Message(WorkCode)

EntryDate = ExcelUtils.getColumnDatas("EntryDate",EnvParams.Opco)
if((EntryDate==null)||(EntryDate=="")){ 
ValidationUtils.verify(false,true,"EntryDate is Needed to CreateAnAccrualJobByJob");
}
Log.Message(EntryDate)

NoForAccrual = ExcelUtils.getColumnDatas("NoForAccrual",EnvParams.Opco)
if((NoForAccrual==null)||(NoForAccrual=="")){ 
ValidationUtils.verify(false,true,"NoForAccrual Number is Needed to CreateAnAccrualJobByJob");
}
Log.Message(NoForAccrual)

}


function Create_New_Job(){ 
  
    //Creation of Job
    
    Ipreparation_ID = TestRunner.testCaseId;
    IpreparationUnit = TestRunner.unitName; 
    TestRunner.TempUnit = IpreparationUnit;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Accural for Job")
    }
    
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Accural for Job")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    JobNo = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    
    if((JobNo=="")||(JobNo==null)){
      
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
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    JobNo = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)  
     
    //Creation of Budget
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var budgetSheet = ExcelUtils.getColumnDatas("Budget sheet",EnvParams.Opco)
    if(budgetSheet==""){ 
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Accural for Job")
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
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Accural for Job")
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
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    PoNumber = ExcelUtils.getRowDatas("PO Number_"+PO_SO,EnvParams.Opco)

    
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

TestRunner.testCaseId = Ipreparation_ID;
TestRunner.unitName = IpreparationUnit;

Log.Message(JobNo)



TestRunner.JiraStat = true;
TestRunner.JiraUpdate = true;

}


//Go To Job from Menu
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());


}

}


ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}



/*
*
This Function Create Accural for Purchase Order

*/
function GoToAccruals() {
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var Accrualtab =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.AccrualTab;
 WorkspaceUtils.waitForObj(Accrualtab);
Accrualtab.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var JobNoTextBox = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.JobSearchField;

JobNoTextBox.setText(JobNo);

var table = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(2000, "Reading Table Data in Job List");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var labels= Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.LabelOneOfOneResult;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
      labels = labels.Child(i);
      break;
    }
  }

  WorkspaceUtils.waitForObj(labels);

  var i=0;
  while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=600)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

// Selecting Particular Job
  aqUtils.Delay(2000, "Reading Table Data in Job List");
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(JobNo)){ 
      table.Keys("[Down]");
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(0).OleValue.toString().trim());

  var closefilter = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.Closefilter;
  closefilter.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(2000, "Waiting for maconomy to load object");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

  var jobAccrualPannel = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel;
  jobAccrualPannel.Click();
  jobAccrualPannel.MouseWheel(-200);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  var showlines =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite.McPlainCheckboxView.ShowLinesCheckBox;
  var includeFullyAccured =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite2.McPlainCheckboxView.inclueFullyAccured;
  Sys.HighlightObject(showlines)
  Sys.HighlightObject(includeFullyAccured)
  //----------De-Select CheckBox-------------
  if(!showlines.getSelection()){ 
  showlines.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showlines.Click();
  ReportUtils.logStep("INFO", "showlines is UnChecked");
    Log.Message("showlines is UnChecked")
    checkmark = true;
  }
  
  if(includeFullyAccured.getSelection()){ 
includeFullyAccured.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  includeFullyAccured.Click();
  ReportUtils.logStep("INFO", "includeFullyAccured is UnChecked");
    checkmark = true;
  }
  
  
  
  var purchaseorderNoFromField = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget
  var purchaseorderNoToField = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite4.PurchaseorderToField;
  
  var purchaseorderlineNoField = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite3.PruchaseOrderFrom;
  
  var workCodeFrom = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite5.WorkCodeField;
  var workCodeTo = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite5.WorkCodeTo;
  
  Sys.HighlightObject(purchaseorderNoFromField);
  Sys.HighlightObject(purchaseorderNoToField);
  Sys.HighlightObject(workCodeFrom);
  Sys.HighlightObject(workCodeTo);
  
  purchaseorderNoFromField.Click();
  WorkspaceUtils.SearchByValue(purchaseorderNoFromField,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),PoNumber,"Purchase Order Number");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  purchaseorderNoToField.Click();
  WorkspaceUtils.SearchByValue(purchaseorderNoToField,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),PoNumber,"Purchase Order Number");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  workCodeFrom.Click();
  WorkspaceUtils.SearchByValue(workCodeFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),WorkCode,"WorkCode");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  workCodeTo.Click();
  WorkspaceUtils.SearchByValue(workCodeTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),WorkCode,"WorkCode");
  
  var savejob =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.savejobButton;
  savejob.Click();
  
  aqUtils.Delay(3000, "Waiting for purchaseOrderTable load");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

  var purchaseOrderTable =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
  var flag=false;
  for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
  
    if((purchaseOrderTable.getItem(v).getText_2(0).OleValue.toString().trim()==WorkCode)&&(purchaseOrderTable.getItem(v).getText_2(6).OleValue.toString().trim()==PoNumber)){ 

      flag=true;
    purchaseOrderTable.Keys("[Tab][Tab][Tab][Tab]");
    aqUtils.Delay(2000);
    purchaseOrderTable.Keys(EntryDate);  
    aqUtils.Delay(5000);
    purchaseOrderTable.Keys("[Tab][Tab][Tab][Tab]");
    aqUtils.Delay(5000);
    var AccuralSelected = purchaseOrderTable.SWTObject("McPlainCheckboxView", "", 4).SWTObject("Button", "");
    Sys.HighlightObject(AccuralSelected)
    if(!AccuralSelected.getSelection()){ 
    AccuralSelected.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    AccuralSelected.Click();
    ReportUtils.logStep("INFO", "Selected for Accural");
    }
    aqUtils.Delay(5000);
    purchaseOrderTable.Keys("[Tab]");
    purchaseOrderTable.Keys(NoForAccrual);  
    aqUtils.Delay(5000);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  var savePOLine = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SavePOLine
  savePOLine.Click();
  aqUtils.Delay(4000);  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}  
  aqUtils.Delay(1000);
 
      break;
      
    }
    else{ 
      purchaseOrderTable.Keys("[Down]");
    }
  }
  
   if(flag){
  ValidationUtils.verify(flag,true,"Purchase Order Line with Work Code is available in system");
  ValidationUtils.verify(true,true,"Batch Accrual is Successful");
  }
  else{
     ValidationUtils.verify(false,true,"Purchase Order Line with Work Code is not available in system");
  ValidationUtils.verify(false,true,"Batch Accrual is not Successful");
  }
    
  if(savejob.isEnabled()){
  var savejob =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.savejobButton;
  savejob.Click();
 }
   var CreateAccruals =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.CreateAccrual;
   Sys.HighlightObject(CreateAccruals)
   CreateAccruals.Click();
 aqUtils.Delay(5000);
 
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PO Used for Accrual Job by Job",EnvParams.Opco,"Data Management",PoNumber)
ExcelUtils.WriteExcelSheet("Job Number Used for Accrual Job by Job",EnvParams.Opco,"Data Management",JobNo)

  }
  



