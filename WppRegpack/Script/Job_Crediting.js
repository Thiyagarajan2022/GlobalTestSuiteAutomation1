//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


/**
 * This script Create Credit Note for Invoice
 * @author  : Muthu Kumar M
 * @version : 1.0
 * Created Date :02/22/2021
*/

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Job Crediting";

//Global Variable
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var EmployeeNumber = "";
var Language = "";
var JobCrediting_JIRA_ID = "";
var JobCrediting_UnitName_JIRA = "";
var Job_Number = "";
var invoice_Number = "";
var I_Budget_Job_No,WriteOff_Job_No,CarryForward_Job_No,I_OnAccount_Job_No,Time_Material_Job_No,I_Preparation_job_No = "";
var I_Budget_Job_Invoice_No,WriteOff_Job_Invoice_No,CarryForward_Job_Invoice_No,I_OnAccount_Job_Invoice_No,Time_Material_Job_Invoice_No,I_Preparation_job_Invoice_No = "";
var Estimatelines = []; 
var Invoice_Type = "";

//Main Function
function Create_Job_Creaditing() {
  
TextUtils.writeLog("Job Crediting Creation Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Job Crediting script
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
sheetName = "CombinedInvoice";


ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";

Approve_Level =[];
ApproveInfo = [];
Estimatelines = []; 
level =0;
EmployeeNumber = "";
JobCrediting_JIRA_ID = "";
JobCrediting_UnitName_JIRA = "";
I_Budget_Job_No,WriteOff_Job_No,CarryForward_Job_No,I_OnAccount_Job_No,Time_Material_Job_No,I_Preparation_job_No = "";
I_Budget_Job_Invoice_No,WriteOff_Job_Invoice_No,CarryForward_Job_Invoice_No,I_OnAccount_Job_Invoice_No,Time_Material_Job_Invoice_No,I_Preparation_job_Invoice_No = "";
Job_Number = "";
invoice_Number = "";
Invoice_Type = "";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);



getDetails();

aqUtils.Delay(5000, Indicator.Text);
goToJobMenuItem();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
gotoInvoicing()
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
goTo_InvoiceHistory();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
if(Invoice_Type.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "On Account Invoice").OleValue.toString().trim())!=-1){ 
  Invoice_OnAccount();
  Submit_Draft_OnAccount();
}else{ 
  Invoice_Selection();
  Submit_Draft_TM();
}

//Submit_Draft();
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



//Getting Data from Datasheets
function getDetails(){ 
  

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  I_Budget_Job_No = ExcelUtils.getRowDatas("Invoice from Budget Job",EnvParams.Opco);
  WriteOff_Job_No = ExcelUtils.getRowDatas("Write Off Invoicing Job",EnvParams.Opco);
  CarryForward_Job_No = ExcelUtils.getRowDatas("Carryforward Invoicing Job",EnvParams.Opco);
  I_OnAccount_Job_No = ExcelUtils.getRowDatas("Invoice OnAccount Job",EnvParams.Opco);
  Time_Material_Job_No = ExcelUtils.getRowDatas("Time & Material Invocing Job",EnvParams.Opco);
  I_Preparation_job_No = ExcelUtils.getRowDatas("Invoice preparation Job",EnvParams.Opco);
  

  I_Budget_Job_Invoice_No = ExcelUtils.getRowDatas("Invoice from Budget No",EnvParams.Opco);
  WriteOff_Job_Invoice_No = ExcelUtils.getRowDatas("Write Off Invoice No",EnvParams.Opco);
  CarryForward_Job_Invoice_No = ExcelUtils.getRowDatas("Carryforward Invoice No",EnvParams.Opco);
  I_OnAccount_Job_Invoice_No = ExcelUtils.getRowDatas("Invoice OnAccount No",EnvParams.Opco);
  Time_Material_Job_Invoice_No = ExcelUtils.getRowDatas("Time & Material Invocing No",EnvParams.Opco);
  I_Preparation_job_Invoice_No = ExcelUtils.getRowDatas("Invoice preparation No",EnvParams.Opco);

  Job_Number = "";
  invoice_Number = "";

  Log.Message(I_Budget_Job_No!="")
  Log.Message(I_Budget_Job_No!=null)
  Log.Message(I_Budget_Job_Invoice_No!="")
  Log.Message(I_Budget_Job_Invoice_No!=null)
  
  if(((I_Budget_Job_No!=null)&&(I_Budget_Job_No!="")) && ((I_Budget_Job_Invoice_No!="")&&(I_Budget_Job_Invoice_No!=null))){
  Job_Number = I_Budget_Job_No;
  invoice_Number = I_Budget_Job_Invoice_No;
  Log.Message("Used Invoice From Budget Job and Invoice No");
  }
  else if((WriteOff_Job_No!="")&&(WriteOff_Job_No!=null) && (WriteOff_Job_Invoice_No!="")&&(WriteOff_Job_Invoice_No!=null)){
  Job_Number = WriteOff_Job_No;
  invoice_Number = WriteOff_Job_Invoice_No;
  Log.Message("Used Invoice Write-Off Job and Invoice No");
  }
  else if((CarryForward_Job_No!="")&&(CarryForward_Job_No!=null) && (CarryForward_Job_Invoice_No!="")&&(CarryForward_Job_Invoice_No!=null)){
  Job_Number = CarryForward_Job_No;
  invoice_Number = CarryForward_Job_Invoice_No;
  Log.Message("Used Invoice CarryForward Job and Invoice No");
  }
  else if((I_OnAccount_Job_No!="")&&(I_OnAccount_Job_No!=null) && (I_OnAccount_Job_Invoice_No!="")&&(I_OnAccount_Job_Invoice_No!=null)){
  Job_Number = I_OnAccount_Job_No;
  invoice_Number = I_OnAccount_Job_Invoice_No;
  Log.Message("Used Invoice On Account Job and Invoice No");
  }
  else if((Time_Material_Job_No!="")&&(Time_Material_Job_No!=null) && (Time_Material_Job_Invoice_No!="")&&(Time_Material_Job_Invoice_No!=null)){
  Job_Number = Time_Material_Job_No;
  invoice_Number = Time_Material_Job_Invoice_No;
  Log.Message("Used Time and Material Invoice Job and Invoice No");
  }
  else if((I_Preparation_job_No!="")&&(I_Preparation_job_No!=null) && (I_Preparation_job_Invoice_No!="")&&(I_Preparation_job_Invoice_No!=null)){
  Job_Number = I_Preparation_job_No;
  invoice_Number = I_Preparation_job_Invoice_No;
  Log.Message("Used Invoice Preparation Job and Invoice No");
  }
  
  Log.Message(I_Preparation_job_No!="")
  Log.Message(I_Preparation_job_No!=null)
  Log.Message(I_Preparation_job_Invoice_No!="")
  Log.Message(I_Preparation_job_Invoice_No!=null)
  Log.Message(I_Preparation_job_No)
  Log.Message(I_Preparation_job_Invoice_No)
  Log.Message(Job_Number)
  Log.Message(invoice_Number)
  
  if((Job_Number==null)||(invoice_Number==null)){
    ExcelUtils.setExcelName(workBook, "Job Crediting", true);
    Job_Number = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco);
    invoice_Number = ExcelUtils.getColumnDatas("Invoice Number",EnvParams.Opco);
  }

  Log.Message(Job_Number)
  Log.Message(invoice_Number)
    ExcelUtils.setExcelName(workBook, "Job Crediting", true);
    EmployeeNumber = ExcelUtils.getColumnDatas("Employee Number",EnvParams.Opco);
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



//Selecting Job for  Job Crediting in Maconomy
function gotoInvoicing(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var allJobs = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
allJobs.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  var table = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
  
  var job = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(Job_Number);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(3000, "Finding Jobs in Maconomy");


  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==Job_Number){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to for Credit Note");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+Job_Number+") is available in maconommy for Credit Note"); 
  closeFilter.Click();
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  var Invoicing = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
}
}



function goTo_InvoiceHistory(){ 
  
var History = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(History);
History.Click();

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
var table = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(table);
aqUtils.Delay(4000, Indicator.Text);

Invoice_Type = "";
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==invoice_Number){ 
      Invoice_Type = table.getItem(v).getText_2(1).OleValue.toString().trim()
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  

  if(flag){ 
    ValidationUtils.verify(true,true,"Invoice is available for the Job")
  }else{ 
    ValidationUtils.verify(true,false,"Invoice No is NOT available for the Job")
  }

}





//Create Credit Note for On Account Invoice Type
function Invoice_OnAccount(){ 
  
var OnAccount = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(OnAccount);
OnAccount.Click();


  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
  TextUtils.writeLog("Moving to Invoice On Account")
var Action = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.GroupToolItemControl;
Sys.HighlightObject(Action);
Action.Click();


aqUtils.Delay(8000, "Clicking Submit");
Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice to Credit Memo").OleValue.toString().trim());

aqUtils.Delay(8000, "Select invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
TextUtils.writeLog("Invoice to Credit Memo is Clicked in Actions");

var InvoiceNumber = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McValuePickerWidget;
WorkspaceUtils.SearchByValue(InvoiceNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice").OleValue.toString().trim(),invoice_Number,"Invoice");
aqUtils.Delay(3000, "Select invoice");

var Type = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
Type.Click();
aqUtils.Delay(8000, "Select invoice");
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "On Account").OleValue.toString().trim(),"Invoice Type")

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Invoice Number and Type is Selected");
  
var Create = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice to Credit Memo").OleValue.toString().trim());
Sys.HighlightObject(Create);
Create.Click();

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
TextUtils.writeLog("Credit Memo is Created");

var Total = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
Sys.HighlightObject(Total);
Total.Click();

Total = Total.getText().OleValue.toString().trim();

if(Total.indexOf("-")==0){ 

var Approve = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.SingleToolItemControl;
 Sys.HighlightObject(Approve);
 Approve.Click();
 TextUtils.writeLog("Credit Memo is Approved in Invoice On Account Tab");
 
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
  
var DraftInvoice = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(DraftInvoice);
DraftInvoice.Click();
 
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Moving to Draft Invoice Tab")
  

}

}




//Create Credit Note for T&M Invoice Type
function Invoice_Selection(){ 
  

var Selection = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
Sys.HighlightObject(Selection);
Selection.Click();

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Moving to Invoice Selection Tab");
  
var Action = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.GroupToolItemControl;
Sys.HighlightObject(Action);
Action.Click();


aqUtils.Delay(8000, "Clicking Submit");
Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice to Credit Memo").OleValue.toString().trim());

aqUtils.Delay(8000, "Select invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
TextUtils.writeLog("Clicked Invoice to Credit Memo");

var InvoiceNumber = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McValuePickerWidget;
WorkspaceUtils.SearchByValue(InvoiceNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice").OleValue.toString().trim(),invoice_Number,"Invoice");
aqUtils.Delay(3000, "Select invoice");


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var Restore_Jobs_Entries = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McPlainCheckboxView.Button;
if(!Restore_Jobs_Entries.getSelection()){ 
Restore_Jobs_Entries.Click();
aqUtils.Delay(5000, "Restoring Job Entries");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
TextUtils.writeLog("Restoring Job Entries is Checked");
}

var Type = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McPopupPickerWidget;
Type.Click();
aqUtils.Delay(9000, "Select invoice");
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "T&M").OleValue.toString().trim(),"Invoice Type")

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Invoice Number and Type is Entered");
var Create = Aliases.Maconomy.Invoice_to_Credit_Memo.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice to Credit Memo").OleValue.toString().trim());
Sys.HighlightObject(Create);
Create.Click();

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
TextUtils.writeLog("Credit Memo is Created")
var Total = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
Sys.HighlightObject(Total);
Total.Click();

Total = Total.getText().OleValue.toString().trim();

if(Total.indexOf("-")==0){ 

var Approve = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.SingleToolItemControl;
 Sys.HighlightObject(Approve);
 Approve.Click();
 
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
TextUtils.writeLog("Credit Memo is Approved in Invoice Selection Tab");

var DraftInvoice = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(DraftInvoice);
DraftInvoice.Click();
 
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Moving to Draft Invoice Tab")
  
}

}




//Submit the created Draft Credit Memo for On Account
function Submit_Draft_OnAccount(){ 
  
aqUtils.Delay(4000, Indicator.Text);
var CloseFilter = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(CloseFilter);
CloseFilter.Click();

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
  
  
  //Changing Description in English
  if((EnvParams.Country.toUpperCase()=="CHINA") && (Language=="English")){
  var DraftEditing = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
  DraftEditing.Click();
  var grid = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  for(var i=0;i<grid.getItemCount()-1;i++){ 
          grid.Keys("[Tab]");
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
              }
          var Des = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
          Des.Click();
          Des.setText("WorkCode "+(i+1));
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


  grid = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(grid)
//  Log.Message(i)
//  Log.Message(grid.getItemCount()-2)
//  Log.Message(i<grid.getItemCount()-2)
  if(i<grid.getItemCount()-2){
  grid.Keys("[Down]");
  }
  }  
  
  }
  
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  
  //Entering Credit Reason
  if(ImageRepository.ImageSet.Forward.Exists()){ 
  
      
      var CreditNote_Reason = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
      Sys.HighlightObject(CreditNote_Reason);
      CreditNote_Reason.Click();
  
      CreditNote_Reason.setText("Invoice Error Correction");
      aqUtils.Delay(4000, Indicator.Text);
      TextUtils.writeLog("Credit Note Reason is Entered");

      
      if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).isVisible())
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
      else
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
      Sys.HighlightObject(Save);
      Save.Click();
      aqUtils.Delay(3000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      aqUtils.Delay(1000, Indicator.Text);
      TextUtils.writeLog("Credit Note Reason is Saved");
      
      ImageRepository.ImageSet.Forward.Click();
  }else{ 
    
  var Sliding_Panel = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "")
//  var Sliding_Panel = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  Sys.HighlightObject(Sliding_Panel);
  Sliding_Panel.Click();
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
     
      var CreditNote_Reason = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
      Sys.HighlightObject(CreditNote_Reason);
      CreditNote_Reason.Click();
  
      CreditNote_Reason.setText("Invoice Error Correction");
      aqUtils.Delay(4000, Indicator.Text);
      TextUtils.writeLog("Credit Note Reason is Entered");

  
      if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).isVisible())
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
      else
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
      
      Sys.HighlightObject(Save);
      Save.Click();
      aqUtils.Delay(3000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      aqUtils.Delay(1000, Indicator.Text);
      TextUtils.writeLog("Credit Note Reason is Saved");
      
      ImageRepository.ImageSet.Forward.Click();
      
      
  }

  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  
  var SubmitDraft;
//  if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.isVisible())
//  SubmitDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2;
// else
//  SubmitDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;;
            
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
  SubmitDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
 else
  SubmitDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);

  WorkspaceUtils.waitForObj(SubmitDraft);
  for(var i=0;i<SubmitDraft.ChildCount;i++){ 
    if((SubmitDraft.Child(i).isVisible())&&(SubmitDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(SubmitDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      SubmitDraft.Child(i).Click();
      TextUtils.writeLog("Credit Memo is Submitted");
      break;
    }
  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var Excl_Tax = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
var grandTotal = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
//var Payment_Terms = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.Composite.McGroupWidget.Composite.McPopupPickerWidget;


//Finding Payment Terms
var break_MainLoop = false;
var ParentAdd = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.Composite.McGroupWidget;
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
var specification = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var q = 0;
QuoteDetails = [];
var InvoiceMPL = "CreditMemo";
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
  Q_total =parseFloat(Q_BillingTotoal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
  //Q_total =Q_total+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
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
//  if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.isVisible())
//  PrintDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2;
// else
//  PrintDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;;

 if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
  PrintDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
 else
  PrintDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  WorkspaceUtils.waitForObj(PrintDraft);
  for(var i=0;i<PrintDraft.ChildCount;i++){ 
    if((PrintDraft.Child(i).isVisible())&&(PrintDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(PrintDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      PrintDraft.Child(i).Click();
      TextUtils.writeLog("Credit Memo Draft is Printed")
      break;
    }
  } 
  
TextUtils.writeLog("Print Draft is Clicked");
    aqUtils.Delay(5000, Indicator.Text);
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
    
ValidationUtils.verify(true,true,"Print Draft Credit Memo is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Draft Credit Note",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf");




var appvBar = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
ImageRepository.ImageSet.Maximize.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

var DraftApproval = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();

aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);
TextUtils.writeLog("Finding Credit Memo Approvers")
//Finding Approvals
var ApproverTable = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
 var y=0;
for(var i=0;i<ApproverTable.getItemCount();i++){   
   var approvers="";
    if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
    approvers = EnvParams.Opco+"*"+Job_Number+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
    Log.Message("Approver level :" +i+ ": " +approvers);
    Approve_Level[y] = approvers;
    y++;
    }
}
ReportUtils.logStep_Screenshot("");

var closeBar = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ImageRepository.ImageSet.Forward.Click();

aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);


}



//Submit the created Draft Credit Memo for T&M
function Submit_Draft_TM(){ 
  
aqUtils.Delay(4000, Indicator.Text);
var CloseFilter = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(CloseFilter);
CloseFilter.Click();

  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, Indicator.Text);
  
  
  //Changing Description in English
  if((EnvParams.Country.toUpperCase()=="CHINA") && (Language=="English")){
  var DraftEditing = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
  DraftEditing.Click();
  var grid = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  for(var i=0;i<grid.getItemCount()-1;i++){ 
          grid.Keys("[Tab]");
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
              }
          var Des = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
          Des.Click();
          Des.setText("WorkCode "+(i+1));
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


  grid = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(grid)
//  Log.Message(i)
//  Log.Message(grid.getItemCount()-2)
//  Log.Message(i<grid.getItemCount()-2)
  if(i<grid.getItemCount()-2){
  grid.Keys("[Down]");
  }
  }  
  
  }
  
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  
  //Entering Credit Reason
  if(ImageRepository.ImageSet.Forward.Exists()){ 
  

      var CreditNote_Reason = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
      Sys.HighlightObject(CreditNote_Reason);
      CreditNote_Reason.Click();
  
      CreditNote_Reason.setText("Invoice Error Correction");
      aqUtils.Delay(4000, Indicator.Text);
      TextUtils.writeLog("Credit Memo Reason is Entered");

  
      if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).isVisible())
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
      else
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
      Sys.HighlightObject(Save);
      Save.Click();
      aqUtils.Delay(3000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      aqUtils.Delay(1000, Indicator.Text);
      TextUtils.writeLog("Credit Memo Reason is Saved");
      
      ImageRepository.ImageSet.Forward.Click();
  }else{ 
var Sliding_Panel = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "")
//  var Sliding_Panel = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  Sys.HighlightObject(Sliding_Panel);
  Sliding_Panel.Click();
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  

      var CreditNote_Reason = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
      Sys.HighlightObject(CreditNote_Reason);
      CreditNote_Reason.Click();
  
      CreditNote_Reason.setText("Invoice Error Correction");
      aqUtils.Delay(4000, Indicator.Text);
      TextUtils.writeLog("Credit Memo Reason is Entered");
      
      if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).isVisible())
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
      else
      var Save = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
      Sys.HighlightObject(Save);
      Save.Click();
      aqUtils.Delay(3000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      aqUtils.Delay(1000, Indicator.Text);
      TextUtils.writeLog("Credit Memo Reason is Saved");
      
      ImageRepository.ImageSet.Forward.Click();
      
      
  }

  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(1000, Indicator.Text);
  
  var SubmitDraft;

  //  if(Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.isVisible())
//  SubmitDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2;
//else
//SubmitDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;

  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
  SubmitDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
 else
  SubmitDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
  
                
  Sys.HighlightObject(SubmitDraft)
  SubmitDraft.HoverMouse();        

  for(var i=0;i<SubmitDraft.ChildCount;i++){ 
    if((SubmitDraft.Child(i).isVisible())&&(SubmitDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit Draft").OleValue.toString().trim())){
      Sys.HighlightObject(SubmitDraft.Child(i));
      SubmitDraft.Child(i).HoverMouse();        
      ReportUtils.logStep_Screenshot("");
      SubmitDraft.Child(i).Click();
      TextUtils.writeLog("Credit Memo is Submitted");
      break;
    }
  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var Excl_Tax = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McTextWidget;
var grandTotal = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
//var Payment_Terms = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.Composite.McGroupWidget.Composite.McPopupPickerWidget;


//Finding Payment Terms
var break_MainLoop = false;
var ParentAdd = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.Composite.McGroupWidget;
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
var specification = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var q = 0;
QuoteDetails = [];
var InvoiceMPL = "CreditMemo";
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
  Log.Message(Q_BillingTotoal.replace(/,/g, ''))
  Q_total =parseFloat(Q_BillingTotoal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
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
//  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
//  PrintDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
// else
//  PrintDraft = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;;

    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
  PrintDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
 else
  PrintDraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  WorkspaceUtils.waitForObj(PrintDraft);
  for(var i=0;i<PrintDraft.ChildCount;i++){ 
    if((PrintDraft.Child(i).isVisible())&&(PrintDraft.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Draft").OleValue.toString().trim())){
      WorkspaceUtils.waitForObj(PrintDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      PrintDraft.Child(i).Click();
      TextUtils.writeLog("Draft Credit Memo is Printed");
      break;
    }
  } 
  
TextUtils.writeLog("Print Draft is Clicked");
    aqUtils.Delay(5000, Indicator.Text);
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
    
ValidationUtils.verify(true,true,"Print Draft Credit Memo is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Draft Credit Note",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf");




var appvBar = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
ImageRepository.ImageSet.Maximize.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);

var DraftApproval = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();

aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);
TextUtils.writeLog("Finding Approvers for Credit Memo");

//Finding Approvals
var ApproverTable = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
 var y=0;
for(var i=0;i<ApproverTable.getItemCount();i++){   
   var approvers="";
    if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
    approvers = EnvParams.Opco+"*"+Job_Number+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
    Log.Message("Approver level :" +i+ ": " +approvers);
    Approve_Level[y] = approvers;
    y++;
    }
}
ReportUtils.logStep_Screenshot("");

var closeBar = Aliases.Maconomy.Job_Crediting.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ImageRepository.ImageSet.Forward.Click();

aqUtils.Delay(4000, Indicator.Text);
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

if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

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
aqUtils.Delay(4000,"Maconomy loading data")

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")


if(ImageRepository.ImageSet.Show_Filter.Exists()){
ImageRepository.ImageSet.Show_Filter.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

var Maconomy_Index = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("Text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Draft Invoices").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
Maconomy_Index = w.FullName;
}
  
var fg = Maconomy_Index.indexOf("PTabFolder");
Log.Message(fg)
Maconomy_Index = Maconomy_Index.substring(0,fg);
Log.Message(Maconomy_Index)
Maconomy_Index = Maconomy_Index.substring(0,Maconomy_Index.lastIndexOf("."));
Log.Message(Maconomy_Index)
Maconomy_Index = Maconomy_Index.substring(0,Maconomy_Index.lastIndexOf("."));
Log.Message(Maconomy_Index)
Maconomy_Index = Maconomy_Index.substring(Maconomy_Index.lastIndexOf(" ")+1,Maconomy_Index.lastIndexOf(")"));
Log.Message(Maconomy_Index)

//Checking the screen with CloseFilter or ShowFilter
//var table = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder;
//WorkspaceUtils.waitForObj(table);
//Sys.HighlightObject(table);
//if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
//  
//if(Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Visible){
//
//}else{ 
//ImageRepository.ImageSet.Show_Filter.Click();
//}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data")

            
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", Maconomy_Index).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var table = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", Maconomy_Index).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
//var firstCell = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(JobNum);
//var closefilter = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;


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

var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", Maconomy_Index).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();



                  


aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000, Indicator.Text);


var Approve;

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

var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("Text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Draft").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
Approve = w;
}
Log.Message(Approve.FullName);
Sys.HighlightObject(Approve)

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
var printInvoice = "";

/*
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
*/

var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("Text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Credit Memo").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
printInvoice = w;
}
Log.Message(printInvoice.FullName);
Sys.HighlightObject(printInvoice);
printInvoice.Click();


//Saving PDF
TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
aqUtils.Delay(9000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Print Job Credit Memo"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Print Job Credit Memo"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Job Credit Memo")!=-1){
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
aqUtils.Delay(4000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Credit Note",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf");

var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(sFolder+SaveTitle+".pdf");
var textobj;
  try{


var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj).OleValue.toString(); 
  var invoiceName = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Credit Note No").OleValue.toString().trim();
  invoiceName = invoiceName.length;
  Log.Message(invoiceName)
  textobj = textobj.substring(textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Note No").OleValue.toString().trim()+" ")+invoiceName+1);
  Log.Message("CreditNote No:"+textobj.substring(0,textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Note Date").OleValue.toString().trim())))
  textobj = textobj.substring(0,textobj.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Note Date").OleValue.toString().trim()));
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Credit Memo Invoice No",EnvParams.Opco,"Data Management",textobj)
  ExcelUtils.WriteExcelSheet("Credit Memo Job",EnvParams.Opco,"Data Management",JobNum)
  TextUtils.writeLog("Client Credit Memo No"+textobj);


}

  ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
  
  
}
}

}
