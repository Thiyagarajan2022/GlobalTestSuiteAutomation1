//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart
Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
 
var sheetName ="JobEstimateMPL";
var STIME = "";
var jobNumber = "";
var Language = "";


//Main Function
function ValidateJobPdf(){ 
TextUtils.writeLog("Job Quote and Client Approved Estimate Creation Started"); 
Indicator.PushText("waiting for window to open");

Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4).Click();

ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Biller",EnvParams.Opco);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}


excelName = EnvParams.path;
workBook = Project.Path+excelName;
STIME = "";
jobNumber = "";

comapany = EnvParams.Opco;
sheetName ="JobEstimateMPL";

try{
getDetails();
goToJobMenuItem();
goToBudget();
validatingWorkEstimate();
PrintJobBudgetMpl();
validateJobBudgetPdf();
}catch(err){ 
  Log.Message(err);
}

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4).Click();
WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){ 
sheetName ="JobEstimateMPL";  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Quote");
  }
  
function goToJobMenuItem(){

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
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

ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu"); 
}


function goToBudget(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var allJobs = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();
  var table = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
  
  var job = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
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
  ReportUtils.logStep("INFO", "Job is listed in table to create Quote");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job is available in maconommy to create Quote"); 
  closeFilter.Click();
var Budget = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(Budget);
Budget.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(5000,"Selecting Show Budget")
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var show_budget = "";
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).isVisible())
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==1)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Show Budget").OleValue.toString().trim())
    show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);


if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).isVisible())
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").Index==1)    
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Show Budget").OleValue.toString().trim())
    show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);


if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible())
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").Index==1)    
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Show Budget").OleValue.toString().trim())
    show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

Log.Message(show_budget.FullName);
    
WorkspaceUtils.waitForObj(show_budget);
show_budget.Keys("Client Approved Estimate"); 
aqUtils.Delay(5000,"Client Approved Estimate")
ValidationUtils.verify(true,true,"Client Approved Estimate is Selected");
TextUtils.writeLog("Client Approved Estimate is Selected"); 
if(Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.Visible==false){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
ImageRepository.ImageSet.Maximize.Click();
}else{
var approverPanel = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
approverPanel.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ImageRepository.ImageSet.Maximize.Click();
}
    
var allApprover = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(allApprover);
allApprover.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var ApproverTable = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
WorkspaceUtils.waitForObj(ApproverTable);
for(var i=0;i<ApproverTable.getItemCount();i++){   
var approvers="";
if(ApproverTable.getItem(i).getText_2(8)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(false,true,"Level "+i+"Is not Approved");
}
}
TextUtils.writeLog("Working Estimate is Approved"); 
var closeBar = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
    }
    
}


function validatingWorkEstimate(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  C_Currency = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(C_Currency);
  C_Currency = C_Currency.getText().OleValue.toString().trim();
  Log.Message(C_Currency);
  var Revesion = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  WorkspaceUtils.waitForObj(Revesion);
  Revesion = Revesion.getText().OleValue.toString().trim();
  Log.Message(Revesion);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
   Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Refresh();
   var FullBudget = "" 
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.ListPO.text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Full Budget").OleValue.toString().trim())
//   FullBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.ListPO; 
if(Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl.text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Full Budget").OleValue.toString().trim())
   FullBudget = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;

//  var FullBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Log.Message(FullBudget.FullName)  ;
WorkspaceUtils.waitForObj(FullBudget);
FullBudget.Click();
aqUtils.Delay(5000,Indicator.Text)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var linestatus = false;
var specification = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(specification);
Estimate = [];
workEstimate = [];
var j=0;

//for(var i=0;i<Grid.getItemCount();i++){ 
//  var workcode = Grid.getItem(i).getText_2(0).OleValue.toString().trim();
//  var description = Grid.getItem(i).getText_2(3).OleValue.toString().trim();
//  var quantity = Grid.getItem(i).getText_2(6).OleValue.toString().trim();
//  var costBase = Grid.getItem(i).getText_2(7).OleValue.toString().trim();
//  var billingPrice = Grid.getItem(i).getText_2(9).OleValue.toString().trim();
//  if((workcode!="")||(description!="")||(quantity!="")||(costBase!="")||(billingPrice!="")){ 
//   workEstimate[j] = workcode+"*"+description+"*"+quantity+"*"+costBase+"*"+billingPrice+"*";
//  if(EnvParams.Country.toUpperCase()=="INDIA"){ 
//  var Ohsn = Grid.getItem(i).getText_2(12).OleValue.toString().trim();
//  var Ihsn = Grid.getItem(i).getText_2(13).OleValue.toString().trim();
//  workEstimate[j] = workEstimate[j]+Ohsn+"*"+Ihsn+"*";
//  }
//   j++;
//  }
//}

var QuoteMPL = "JobEstimateMPL";
ExcelUtils.setExcelName(workBook,QuoteMPL, true);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Revision",QuoteMPL,Revesion);


var q = 0;
QuoteDetails = [];
for(var i=0;i<specification.getItemCount();i++){ 

  var Q_Desp = specification.getItem(i).getText_2(3).OleValue.toString().trim();
  if((Q_Desp!="")&&(Q_Desp!=null)){
  var Q_Qty = specification.getItem(i).getText_2(6).OleValue.toString().trim();
  var Q_Billing = specification.getItem(i).getText_2(9).OleValue.toString().trim();
  var Q_BillingTotoal = specification.getItem(i).getText_2(10).OleValue.toString().trim();
  var Q_Tax1 = specification.getItem(i).getText_2(15).OleValue.toString().trim();
  var Q_Tax2 = specification.getItem(i).getText_2(16).OleValue.toString().trim();
if(EnvParams.Country.toUpperCase()=="INDIA"){
  var HSN = specification.getItem(i).getText_2(12).OleValue.toString().trim();
  }
//  var Q_Tax2currency = specification.getItem(i).getText_2(10).OleValue.toString().trim();
//  var Q_total = parseFloat(Q_BillingTotoal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
//  QuoteDetails[q] = Q_Desp+"*"+Q_Qty+"*"+Q_Billing+"*"+Q_BillingTotoal+"*"+Q_Tax1+"*"+Q_Tax2+"*"+Q_Tax1currency+"*"+Q_Tax2currency+"*"+Q_total.toFixed(2)+"*";
  Log.Message(QuoteDetails[q]);
  q++;
  ExcelUtils.setExcelName(workBook,QuoteMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,QuoteMPL,Q_Desp);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,QuoteMPL,Q_Qty);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,QuoteMPL,Q_Billing);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TotalBilling_"+q,QuoteMPL,Q_BillingTotoal);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1_"+q,QuoteMPL,Q_Tax1);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2_"+q,QuoteMPL,Q_Tax2);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"HSN_"+q,QuoteMPL,HSN);
//  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2currency_"+q,QuoteMPL,Q_Tax2currency);
//  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Total_"+q,QuoteMPL,Q_total);

  }
  }
  q++;
Log.Message(q)
for(var i=q;i<11;i++){ 
  ExcelUtils.setExcelName(workBook,QuoteMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TotalBilling_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"HSN_"+q,QuoteMPL,"");
//  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2currency_"+q,QuoteMPL,"");
//  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Total_"+q,QuoteMPL,"");
}
var total = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.Composite.Composite.McTextWidget;
total = total.getText().OleValue.toString().trim();
ExcelUtils.setExcelName(workBook,QuoteMPL, true);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"GrandTotal",QuoteMPL,total);
}

function PrintJobBudgetMpl(){ 
  
if(Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.isVisible())
  var print = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.SingleToolItemControl;
  else
  var print = Aliases.Maconomy.JobBudgetMPL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
  
  WorkspaceUtils.waitForObj(print);
  print.Click();
  TextUtils.writeLog("Print Job Budget is Clicked and saved"); 
aqUtils.Delay(5000, Indicator.Text);
    
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_JobBudget"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_JobBudget"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("P_JobBudget")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
WorkspaceUtils.waitForObj(pdf);
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
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print Job Budget is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Job Budget",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

}


function formatMoney(amount, decimalCount = 2, decimal = ".", thousands = ",") {
  try {
    decimalCount = Math.abs(decimalCount);
    decimalCount = isNaN(decimalCount) ? 2 : decimalCount;

    const negativeSign = amount < 0 ? "-" : "";

    let i = parseInt(amount = Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
    let j = (i.length > 3) ? i.length % 3 : 0;

    return negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");
  } catch (e) {
    console.log(e)
  }
};


function validateJobBudgetPdf()
{
  var fileName = "";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("PDF Job Budget",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"Job Budget PDF is needed to validate");
  }
  
  var docObj;
  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName);
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
  ExcelUtils.setExcelName(workBook, "Data Management", true);
 
  var pdflineSplit = docObj.split("\r\n");
  
 
  var street = ReadExcelSheet("Street 1",EnvParams.Opco,"CreateClient");
  var postCode = ReadExcelSheet("Post Code",EnvParams.Opco,"CreateClient");
  var postDistrict = ReadExcelSheet("Post District",EnvParams.Opco,"CreateClient");
  var country = ReadExcelSheet("Country",EnvParams.Opco,"CreateClient");
  var Attn = ReadExcelSheet("Attn.",EnvParams.Opco,"CreateClient");
  var TaxNo = ReadExcelSheet("Tax No.",EnvParams.Opco,"CreateClient");
  

ExcelUtils.setExcelName(workBook, "Data Management", true);
var clientName = ExcelUtils.getRowDatas("Global Client Name",EnvParams.Opco)
if((clientName=="")||(clientName==null)){
clientName = ReadExcelSheet("Client Name",EnvParams.Opco,"CreateClient");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
var productName = ExcelUtils.getRowDatas("Global Product Name",EnvParams.Opco)
if((productName=="")||(productName==null)){
productName = ReadExcelSheet("Product Name",EnvParams.Opco,"CreateClient");
}

   if((EnvParams.Country.toUpperCase()=="INDIA")|| (EnvParams.Country.toUpperCase()=="CHINA") || (EnvParams.Country.toUpperCase()=="HONG KONG"))
   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget").OleValue.toString().trim());
   else if((EnvParams.Country.toUpperCase()=="SPAIN") || (EnvParams.Country.toUpperCase()=="MALAYSIA") || (EnvParams.Country.toUpperCase()=="UAE") || (EnvParams.Country.toUpperCase()=="EGYPT") || (EnvParams.Country.toUpperCase()=="QATAR"))
   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "QUOTE").OleValue.toString().trim());
   else if(EnvParams.Country.toUpperCase()=="SINGAPORE")
   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "QUOTATION").OleValue.toString().trim());
    if(index>=0){
          ReportUtils.logStep("INFO","Heading is available in Pdf")
          ValidationUtils.verify(true,true,"Heading is available in Pdf")
          TextUtils.writeLog("Heading is available in Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Heading is not available in Pdf")
   var index = pdflineSplit.indexOf(clientName);        
    if(index>=0){
          ReportUtils.logStep("INFO",clientName+"ClientName is matching with Pdf")
          ValidationUtils.verify(true,true,clientName+" ClientName is matching with Pdf")
          TextUtils.writeLog(clientName+" ClientName is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"ClientName is not same in pdf");
          if(EnvParams.Country.toUpperCase()=="INDIA"){
   var index = pdflineSplit.indexOf(productName);
    if(index>=0){
          ReportUtils.logStep("INFO",productName+" ProductName is matching with Pdf")
          ValidationUtils.verify(true,true,productName+" ProductName is matching with Pdf")
          TextUtils.writeLog(productName+" ProductName is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"ProductName is not same in pdf"); 
          }
                
   var index = pdflineSplit.indexOf(street);
    if(index>=0){
          ReportUtils.logStep("INFO",street+" Street is matching with Pdf")
          ValidationUtils.verify(true,true,street+" Street is matching with Pdf")
          TextUtils.writeLog(street+" Street is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Street is not same in pdf");
   var index = pdflineSplit.indexOf(postCode+"  "+postDistrict);
    if (index == -1)
        index = pdflineSplit.indexOf(postCode+" "+postDistrict);
    if(index>=0){
          ReportUtils.logStep("INFO",postCode+" "+postDistrict+" PostCode and Post District is matching with Pdf")
          ValidationUtils.verify(true,true,postCode+" "+postDistrict+" PostCode and Post District are matching with Pdf")
          TextUtils.writeLog(postCode+" "+postDistrict+" PostCode and Post District are matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"PostCode and Post District are not same in pdf");
   var index = pdflineSplit.indexOf(country);
    if(index>=0){
          ReportUtils.logStep("INFO",country+" Country is matching with Pdf")
          ValidationUtils.verify(true,true,country+" Country is matching with Pdf")
          TextUtils.writeLog(country+" Country is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Country is not same in pdf");
   
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Validate pdf");
   
  var j, x, pdfJobNum, pointer, pdfJobName;
  
  var jobName = ReadExcelSheet("Job_name",EnvParams.Opco,"JobCreation");
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  var JobCurrency = ExcelUtils.getRowDatas("Client Currency",EnvParams.Opco)
  if((JobCurrency=="")||(JobCurrency==null)){
  JobCurrency = ReadExcelSheet("Currency",EnvParams.Opco,"CreateClient");
  }
  var productNumber = ReadExcelSheet("Global Product Number",EnvParams.Opco,"Data Management");
  var clientNumber = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
  
  
  var sheetName = "JobEstimateMPL";
  ExcelUtils.setExcelName(workBook, sheetName, true);

  var revision = ExcelUtils.getColumnDatas("Revision",EnvParams.Opco); 
  
  
 Log.Message(pdflineSplit.length)
  for (j=0; j<pdflineSplit.length; j++)
  {
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Attention").OleValue.toString().trim()))
    {
        if(!pdflineSplit[j].includes(Attn))
        ValidationUtils.verify(false,true,"Attention is not same in pdf");
        else{
          ReportUtils.logStep("INFO",Attn+" Attention is matching with Pdf");
        ValidationUtils.verify(true,true,Attn+" Attention is matching with Pdf");
        TextUtils.writeLog(Attn+" Attention is matching with Pdf");
        }
      }

    if(pdflineSplit[j].includes("Version No"))
    {
      x= pdflineSplit[j].split(":");
      pdfVersionNo = x[1].trim();
       if(pdfVersionNo!=revision)
        ValidationUtils.verify(false,true,"Version Number is not same in pdf");
        else{
        ReportUtils.logStep("INFO",revision+" Version Number is matching with Pdf")
        ValidationUtils.verify(true,true,revision+" Revesion Number is matching with Pdf")
        TextUtils.writeLog(revision+" Version Number is matching with Pdf")
        }
    }
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job No").OleValue.toString().trim()))
    {
       if(!pdflineSplit[j].includes(jobNumber))
        ValidationUtils.verify(false,true,"Job Number is not same in pdf");
        else{
        ReportUtils.logStep("INFO",jobNumber+" Job Number is matching with Pdf")
        ValidationUtils.verify(true,true,jobNumber+" Job Number is matching with Pdf")
        TextUtils.writeLog(jobNumber+" Job Number is matching with Pdf")
        }
    }
     if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Name").OleValue.toString().trim()))
    {
     
      if(pdflineSplit[j].includes(jobName)){
          ReportUtils.logStep("INFO",jobName+" Job Name is matching with Pdf")
          ValidationUtils.verify(true,true,jobName+" Job Name is matching with Pdf")
          TextUtils.writeLog(jobName+" Job Name is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Job Name is not same in Quote");
    }
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Currency").OleValue.toString().trim()))
    {
        if(pdflineSplit[j].includes(JobCurrency))
         {
          ReportUtils.logStep("INFO","Job Currency is matching with Pdf")
          ValidationUtils.verify(true,true,JobCurrency+" Job Currency is matching with Pdf")
          TextUtils.writeLog(JobCurrency+" Job Currency is matching with Pdf")
          }
          else{
          ValidationUtils.verify(false,true,"Job Currency is not same in Quote");
        }
    }

     if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Product Number").OleValue.toString().trim()))
    {
        if(pdflineSplit[j].includes(pdfproductNumber))
         {
          ReportUtils.logStep("INFO",pdfproductNumber+" Product Number is matching with Pdf")
          ValidationUtils.verify(true,true,pdfproductNumber+" Product Number is matching with Pdf")
          TextUtils.writeLog(pdfproductNumber+" Product Number is matching with Pdf")
          }
          else{
          ValidationUtils.verify(false,true,"Product Number is not same in pdf");
        }
    }
   if(pdflineSplit[j].includes("Product Name"))
    {
       if(pdflineSplit[j].includes(pdfproductName))
         {
          ReportUtils.logStep("INFO",productName+" Product Name is matching with Pdf")
          ValidationUtils.verify(true,true,productName+" Product Name is matching with Pdf")
          TextUtils.writeLog(productName+" Product Name is matching with Pdf")
          }
          else{
          ValidationUtils.verify(false,true,"Product Name is not same in pdf");
        }
    }
   if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client No").OleValue.toString().trim()))
    {      
       if(pdflineSplit[j].includes(clientNumber))
         {
          ReportUtils.logStep("INFO",clientNumber+"Client Number is matching with Pdf")
          ValidationUtils.verify(true,true,clientNumber+" Client Number is matching with Pdf")
          TextUtils.writeLog(clientNumber+" Client Number is matching with Pdf")
          }
          else{
          ValidationUtils.verify(false,true,"Client Number is not same in pdf");
        }
    }  
  }
  
   if(EnvParams.Country.toUpperCase()=="INDIA"){
   
    var clientGST = ReadExcelSheet("Tax No.",EnvParams.Opco,"CreateClient");
    var pos = ReadExcelSheet("State Code",EnvParams.Opco,"CreateClient");
    var pdfClientGST, pdfPOS;     
    pointer = pdflineSplit.indexOf("Client GST Details");  // Start searching for client GST details from this Section 
       if(pointer>=0){  
           for (j=pointer; j<pdflineSplit.length; j++)
          {
            Log.Message(pdflineSplit[j])
             if(pdflineSplit[j].includes("GSTIN"))
              {
                x= pdflineSplit[j].split(":");
                pdfClientGST = x[1].trim();
               if(clientGST!=pdfClientGST)
                ValidationUtils.verify(false,true,"clientGST is not same in pdf");
               else
               {
                ReportUtils.logStep("INFO",clientGST+" clientGST is matching with Pdf")
                ValidationUtils.verify(true,true,clientGST+" clientGST is matching with Pdf")
                TextUtils.writeLog(clientGST+" clientGST is matching with Pdf")
               }
             }
             if(pdflineSplit[j].includes("Place of Supply"))
              {               
               if(pdflineSplit[j].includes(pos))
               {
                ReportUtils.logStep("INFO",pos+" POS is matching with Pdf")
                ValidationUtils.verify(true,true,pos+" POS is matching with Pdf")
                TextUtils.writeLog(pos+" POS is matching with Pdf")
                break;
                }
               else
               {
                ValidationUtils.verify(false,true,"POS is not same in pdf");
                break;
               }
            }
       } 
      }  
    pointer =-1;   // Setting again pointer to 1
    pointer = pdflineSplit.indexOf("Agency GST Details");
    
    if(pointer>=0){
    var pdfPan, pdfGstin, pdfCin,pdfStatePOS;  
    var pan = ReadExcelSheet("OpCo PAN",EnvParams.Opco,"OpCo Details");
    var gstin = ReadExcelSheet("OpCo GSTIN",EnvParams.Opco,"OpCo Details");
    var cin = ReadExcelSheet("CIN/UIN",EnvParams.Opco,"OpCo Details");
    var statePOS = ReadExcelSheet("OpCo Company POS",EnvParams.Opco,"OpCo Details");
    
      for (j=pointer; j<=pdflineSplit.length; j++)
      {
      if(pdflineSplit[j].includes("PAN"))
      {
        x= pdflineSplit[j].split(":");
        pdfPan = x[1].trim();
         if(pan!=pdfPan)
          ValidationUtils.verify(false,true,"PAN is not same in pdf");
         else{
          ReportUtils.logStep("INFO",pan+" PAN is matching with Pdf")
          ValidationUtils.verify(true,true,pan+" PAN is matching with Pdf")
          TextUtils.writeLog(pan+" PAN is matching with Pdf")
          }
      }
       if(pdflineSplit[j].includes("GSTIN"))
      {
        x= pdflineSplit[j].split(":");
        pdfGstin = x[1].trim();
        if(gstin!=pdfGstin)
            ValidationUtils.verify(false,true,"GSTIN is not same in pdf");
        else{
            ReportUtils.logStep("INFO",gstin+" GSTIN is matching with Pdf")
          ValidationUtils.verify(true,true,gstin+" GSTIN is matching with Pdf")
          TextUtils.writeLog(gstin+" GSTIN is matching with Pdf")
          }
      }
     
      if(pdflineSplit[j].includes("State"))
      {
        if(pdflineSplit[j].includes(statePOS))
        {
          ReportUtils.logStep("INFO",cin+" State is matching with Pdf")
          ValidationUtils.verify(true,true,cin+" State is matching with Pdf")
          TextUtils.writeLog(cin+" State is matching with Pdf")
          }
        else
          ValidationUtils.verify(false,true,"State is not same in pdf");
          
       }
        if(pdflineSplit[j].includes("CIN/UIN"))
      {
        x= pdflineSplit[j].split(":");
        pdfCin = x[1].trim();
        if(cin!=pdfCin)
        {
            ValidationUtils.verify(false,true,"CIN/UIN is not same in pdf");
            break;
            }
        else{
            ReportUtils.logStep("INFO",cin+" CIN/UIN is matching with Pdf")
          ValidationUtils.verify(true,true,cin+" CIN/UIN is matching with Pdf")
          TextUtils.writeLog(cin+" CIN/UIN is matching with Pdf")
          break;
          }
      }
      }
      }
    else
        ValidationUtils.verify(false,true,"Agency GST Details Section is not displayed in pdf");
    }
  
  var sheetName = "JobEstimateMPL";
  ExcelUtils.setExcelName(workBook, sheetName, true);
  Log.Message(workBook)
  Log.Message(sheetName)
  for(var i=1;i<11;i++){
  var temp = "";
  var desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco);
  //Log.Message(desp)
  if(desp!=""){
  temp = temp + desp+" ";
  
   if(EnvParams.Country.toUpperCase()=="INDIA"){
  var hsnCode = ExcelUtils.getColumnDatas("HSN_"+i,EnvParams.Opco);
  if(hsnCode!=""){
  temp = temp + "HSN Code: "+ hsnCode+" ";
  }}
  
  var qty = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco);
  if(qty!=""){
  temp = temp + qty+" ";
  }
  var unitPrice = ExcelUtils.getColumnDatas("UnitPrice_"+i,EnvParams.Opco);
  if(unitPrice!="")
    temp = temp + unitPrice+" ";
  
  var billingTotal = ExcelUtils.getColumnDatas("TotalBilling_"+i,EnvParams.Opco);
  if(billingTotal!="")
    temp = temp + billingTotal+" ";

   if(EnvParams.Country.toUpperCase()=="INDIA"){ 
        var rate = ExcelUtils.getColumnDatas("Tax1_"+i,EnvParams.Opco);
//        temp = temp + "0.00 ";   
        var matches = rate.match(/(\d+)/); 
        if (matches) 
        temp = temp + matches[1]+".00 "; 
    }
  
  Log.Message("From Excel :"+temp.trim()) 
  var found = false;
  temp = temp.trim();
   for (z=10; z<pdflineSplit.length; z++) // z= 10 Excluded first few lines as it contains Client information
   {
      if(pdflineSplit[z].includes(temp.trim())){
         ReportUtils.logStep("INFO",temp+" is matching with Pdf")
          ValidationUtils.verify(true,true,temp+" Matched with pdf"); 
          TextUtils.writeLog(temp+" Matched with pdf"); 
        found = true;
        break;
      }
      if(z==pdflineSplit.length-1 && !found){
        ValidationUtils.verify(false,true,temp+" is not matching with the Pdf"); 
        break;
      }
   } 
   }
   else
    break;
  }
  
    var grandTotal = ExcelUtils.getColumnDatas("GrandTotal",EnvParams.Opco);
    found = false;
    for (var j=20; j<=pdflineSplit.length; j++)
      {
        if(pdflineSplit[j].includes(JobCurrency+" "+grandTotal));
        {
          ReportUtils.logStep("INFO",grandTotal+" Total is matching with Pdf");
          ValidationUtils.verify(true,true,grandTotal+" Total is matching with Pdf")
          TextUtils.writeLog(grandTotal+" Total is matching with Pdf")
          found = true;
          break;
          }
          if(j==pdflineSplit.length-1 && !found){
          ValidationUtils.verify(false,true,"Total is not found in pdf");
          break;
      }   
      }
    
}


function getTextFromPDF(docObj){
 var textobj;
  try{
  obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj);
  Log.Message(textobj)
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  return textobj;
}



