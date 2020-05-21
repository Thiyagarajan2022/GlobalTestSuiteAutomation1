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
var sheetName = "CreatePurchaseOrder";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice = "";
var PONumber = "";
var workcodeList = [];
 
function CreatePurchaseOrder(){ 
Indicator.PushText("waiting for window to open");
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);
  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreatePurchaseOrder";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice = "";
PONumber = "";
Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);


getDetails();
ExcelUtils.setExcelName(workBook, sheetName, true);
  gotoMenu();
//  Delay(5000);
  selectJobs();
  listPurchaseOrder();
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
closeAllWorkspaces();
}

function getDetails(){ 
  
ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorID = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
if((VendorID=="")||(VendorID==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorID = ExcelUtils.getColumnDatas("Vendor Number",EnvParams.Opco)
}
if((VendorID==null)||(VendorID=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Purchase Order");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Job_Number = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
if((Job_Number=="")||(Job_Number==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_Number = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
}
if((Job_Number==null)||(Job_Number=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to Create a Purchase Order");
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
//aqUtils.Delay(3000, Indicator.Text);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|Jobs");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Jobs");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Moved to Jobs from Jobs Menu");
}


function selectJobs(){ 
  var allJobs = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.AllJobs;
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();
  ReportUtils.logStep_Screenshot("");
//  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid;
  var firstcell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid.CompanyNumber;
  var closeFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.CloseFilterList;
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  WorkspaceUtils.waitForObj(firstcell);
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid.Jobno;
  WorkspaceUtils.waitForObj(job);
  job.Click();
  job.setText(Job_Number);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(3000, Indicator.Text);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==Job_Number){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to create PO");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job is listed in table to create PO");
  closeFilter.Click();
//  aqUtils.Delay(8000, Indicator.Text);
  var jobActivities = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.JobActivities;
  WorkspaceUtils.waitForObj(jobActivities);
  jobActivities.Click();
  var iniatePOfromBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.InitiatePO;
  WorkspaceUtils.waitForObj(iniatePOfromBudget);
  iniatePOfromBudget.Click();
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Initiate Purchase Order is clicked");
//  aqUtils.Delay(5000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.POTable.McGrid;
  WorkspaceUtils.waitForObj(table);

var jB = true;
var StartPO = false;

for(var i=1;i<=10;i++){
sheetName = "JobBudgetCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
//Log.Message("wCodeID :"+wCodeID)
if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
 jB = false; 
 break;
}
}

if(jB){
StartPO = true;
sheetName = "CreatePurchaseOrder";
}

  
  
  
  
  
  
  
  
  
Log.Message(sheetName);
  
   for(var i=1;i<=10;i++){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
  var ChkWorkCode = false;
  if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
    for(var j=0;j<table.getItemCount();j++){
    if(table.getItem(j).getText_2(3).OleValue.toString().trim()==wCodeID){ 
      ChkWorkCode = true;
      break;
    }
    }  
    if(!ChkWorkCode)
    ValidationUtils.verify(false,true,"Given WorkCode "+wCodeID+" is not availble in Maconomy Screen");
  }
  }
  workcodeList = [];
  var wcL = 0;
  for(var j=0;j<table.getItemCount();j++){
       for(var i=1;i<=10;i++){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
  if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
  if(table.getItem(j).getText_2(3).OleValue.toString().trim()==wCodeID){ 
  workcodeList[wcL] = table.getItem(j).getText_2(3).OleValue.toString().trim()+"*"+table.getItem(j).getText_2(4).OleValue.toString().trim();
  wcL++;
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  var selected = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.POTable.McGrid.McPlainCheckboxView.Selected;
  if(!selected.getSelection()){ 
  selected.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  selected.Click();
  }
  aqUtils.Delay(1000, Indicator.Text);
  // 8 TAB to move to Vendor
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  var vendor = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.POTable.McGrid.vendor;
  vendor.Click();
  WorkspaceUtils.SearchByValues_Col_1_all(vendor,"Vendor",VendorID,"Vendor Number","All Vendors");
  aqUtils.Delay(1000, Indicator.Text);
  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.RemarksSave;
  save.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  save.Click();
  aqUtils.Delay(3000, "Saving Changes");
  // 9 SHIFT+TAB to move to LineType
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  }
  }
  }
if(j<table.getItemCount()-2)
    table.Keys("[Down]");


  }
  
  }
//  aqUtils.Delay(4000,Indicator.Text);
  var createPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.CreatePO;
  WorkspaceUtils.waitForObj(createPurchaseOrder);
  createPurchaseOrder.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  createPurchaseOrder.Click();
//  aqUtils.Delay(8000,Indicator.Text);
  
  var label = Sys.Process("Maconomy").SWTObject("Shell", "Jobs").SWTObject("Label", "*");
  ReportUtils.logStep(label.getText());
  Log.Message(label.getText())
  PONumber = label.WndCaption;
  PONumber = PONumber.substring(PONumber.lastIndexOf(" ")+1);
  ValidationUtils.verify(true,true,"Purchase Order Number :"+PONumber);   
  var Okay = Sys.Process("Maconomy").SWTObject("Shell", "Jobs").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  Okay.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  Okay.Click();
  aqUtils.Delay(2000,"Created PO Number :"+PONumber);
  TextUtils.writeLog("Purchase Order is Created :"+PONumber);
}


function listPurchaseOrder(){ 
  var listPO = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.ListPO;
  WorkspaceUtils.waitForObj(listPO);
  listPO.Click();
//  aqUtils.Delay(5000,Indicator.Text);
  var myPO = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.MyOpenPO;
  WorkspaceUtils.waitForObj(myPO);
  myPO.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid;
  var firstcell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid.companyNo;
  var closeFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  
  var POcolumn = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid.POcolumn;
  POcolumn.Click();
  POcolumn.setText(PONumber);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(3000, "Reading table data");
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==PONumber){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
    closeFilter.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    closeFilter.Click();
    TextUtils.writeLog("Created Purchase order is listed in Table");
//    aqUtils.Delay(5000, Indicator.Text);
    
//    var POLine = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
//    POLine.Click();
//    aqUtils.Delay(5000, Indicator.Text);
    
  var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-40);
  aqUtils.Delay(3000, "Reading Vendor Currency");
var ClientCurrency =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McTextWidget;
//WorkspaceUtils.waitForObj(ClientCurrency);
ClientCurrency =  ClientCurrency.getText();  
Log.Message(ClientCurrency);
ExcelUtils.setExcelName(workBook, "CountryCurrency", true);
var ContryCurrency = ExcelUtils.getRowDatas(EnvParams.Country,"Currency");
Log.Message(ContryCurrency)
var ExchangeRate;
var BaseCurrency;
  ExcelUtils.setExcelName(workBook, "ExchangeRate", true);
  if(ContryCurrency!="GBP")  
  ExchangeRate = ExcelUtils.getRowDatas(ContryCurrency,"Exchange Rate");
  else
  ExchangeRate = "1.00";
  if(ClientCurrency!=ContryCurrency)  
  BaseCurrency = ExcelUtils.getRowDatas(ClientCurrency,"Exchange Rate");
  else
  BaseCurrency = "1.00";
  Log.Message("ExchangeRate :"+ExchangeRate);
  Log.Message("BaseCurrency :"+BaseCurrency)  
    
    
    var POLine = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
    WorkspaceUtils.waitForObj(POLine);
    POLine.Click();
//    aqUtils.Delay(5000, Indicator.Text);    
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
  WorkspaceUtils.waitForObj(table);
  for(var j=0;j<table.getItemCount();j++){
   for(var i=1;i<=10;i++){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
  var Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
  var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
  var UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
  var OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
  var IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
  

  ExcelUtils.setExcelName(workBook, "CreatePurchaseOrder", true);
  var POS = ExcelUtils.getColumnDatas("POS",EnvParams.Opco)

  
//  Log.Message(wCodeID)
//  Log.Message(Desp)
//  Log.Message(Qly)
//  Log.Message(UnitPrice)
//  Log.Message(OHSN)
//  Log.Message(IHSN)
//  Log.Message(POS)
  
  for(var wcL=0;wcL<workcodeList.length;wcL++){
  if((wCodeID!="")&&(wCodeID!=null)&&(workcodeList[wcL].indexOf(wCodeID)!=-1)){
  var temp = workcodeList[wcL].split("*");
  if(table.getItem(j).getText_2(3).OleValue.toString().trim()==temp[1]){ 
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
//Delay(2000);
var jobNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
jobNo.Keys("[Tab][Tab]");

var workcode = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  workcode.Click();
  aqUtils.Delay(1000, Indicator.Text);
//  WorkspaceUtils.SearchByValue(workcode,"Work Code",wCodeID,"WorkCode");
  workcode.Keys("[Tab]");
var detailedDescription = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if((Desp!="")&&(detailedDescription.getText()!=Desp)){
   detailedDescription.setText(Desp);
   ReportUtils.logStep_Screenshot("");
   }else{ 
   ValidationUtils.verify(false,true,"Detailed Description Needed to create PurchaseOrder");
     }
   detailedDescription.Keys("[Tab]"); 
var Quantity = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
   if((Qly!="")&&(Quantity.getText()!=Qly)){
   Quantity.setText(Qly);
   ReportUtils.logStep_Screenshot("");
   }
   Quantity.Keys("[Tab]");
var Unit_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
   if((UnitPrice!="")&&(Unit_Price.getText()!=UnitPrice)){
   Unit_Price.setText(UnitPrice);
   ReportUtils.logStep_Screenshot("");
     }
     
//Log.Message("OHSN :"+OHSN);
//Log.Message("IHSN :"+IHSN);
//Log.Message("POS :"+POS);
  if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_PurchaseOrderfromJobActivities.IND_Specific",Unit_Price,OHSN,IHSN,POS);
   
//  UnitPrice.Keys("[Tab][Tab][Tab]");
//aqUtils.Delay(2000, Indicator.Text);
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(save);
if(save.toolTipText.OleValue.toString().trim().indexOf("Save Purchase Order Line")!=-1){
save.HoverMouse();
ReportUtils.logStep_Screenshot("");
save.Click();
TextUtils.writeLog("Modified Lines are Saved");
}
//Delay(5000);

  aqUtils.Delay(3000, "Saving Changes");
  
  
  // 6 SHIFT+TAB to move to LineType
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  
if(EnvParams.Country.toUpperCase()=="INDIA"){
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  }
 

aqUtils.Delay(4000, "Validating Tax");
  var tableGrid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var currency_Amount = tableGrid.getItem(j).getText_2(7).OleValue.toString().trim();
  var local_currency_Amount = tableGrid.getItem(j).getText_2(8).OleValue.toString().trim();
  var Taxcode1 = tableGrid.getItem(j).getText_2(12).OleValue.toString().trim();
  var Taxcode2 = tableGrid.getItem(j).getText_2(13).OleValue.toString().trim();
  var Tax_Amount_currency_1 = tableGrid.getItem(j).getText_2(15).OleValue.toString().trim();
  var Tax_Amount_currency_2 = tableGrid.getItem(j).getText_2(17).OleValue.toString().trim();
  var Tax_Amount_1_base = tableGrid.getItem(j).getText_2(14).OleValue.toString().trim();
  var Tax_Amount_2_base = tableGrid.getItem(j).getText_2(16).OleValue.toString().trim();
  var Tax_Amount = tableGrid.getItem(j).getText_2(18).OleValue.toString().trim();
  currency_Amount = currency_Amount.replace(/,/g, '');
  local_currency_Amount = local_currency_Amount.replace(/,/g, '');
  Tax_Amount_currency_1 = Tax_Amount_currency_1.replace(/,/g, '');
  Tax_Amount_currency_2 = Tax_Amount_currency_2.replace(/,/g, '');
  Tax_Amount_1_base = Tax_Amount_1_base.replace(/,/g, '');
  Tax_Amount_2_base = Tax_Amount_2_base.replace(/,/g, '');
  Tax_Amount = Tax_Amount.replace(/,/g, '');
  
  var CA = parseFloat(Qly)*parseFloat(UnitPrice);
  CA = CA.toFixed(2);
  
  var convertCurr;
  var lcA;
  
  if(ClientCurrency==ContryCurrency){ 
//    Log.Message(Qly)
//    Log.Message(UnitPrice)
    lcA = parseFloat(Qly)*parseFloat(UnitPrice);
    Log.Message(lcA)
  }
  else if(ClientCurrency!="GBP"){
   convertCurr = 1/BaseCurrency;
//     Log.Message("convertCurr :"+convertCurr)
  var QtyXCurr = parseFloat(convertCurr)*parseFloat(CA);
//  Log.Message("QtyXCurr :"+QtyXCurr)
   lcA = parseFloat(QtyXCurr)*parseFloat(ExchangeRate);
//  Log.Message("lcA :"+lcA)
  }
  else{ 
    lcA = parseFloat(CA)*parseFloat(ExchangeRate);
  }
//  Log.Message(lcA)
  lcA = lcA.toFixed(2);
//  Log.Message(lcA)
  var lowerRange = parseFloat(lcA)-parseFloat("1000.00");
  var higherRange = parseFloat(lcA)+parseFloat("1000.00");

//  Log.Message(Taxcode1);
//  Log.Message(Taxcode2);
  if((Taxcode1=="")&&(Taxcode2==""))
  ValidationUtils.verify(false,true,"Tax Code 1 and Tax Code 2 is not Populated");
  if(Taxcode1!="")
  ValidationUtils.verify(true,true,"Tax Code 1 is populated");
  if(Taxcode2!="")
  ValidationUtils.verify(true,true,"Tax Code 2 is populated");
  
  
//Log.Message(lowerRange) 
//Log.Message(higherRange) 
//Log.Message(local_currency_Amount)

  if(CA==currency_Amount)
  ValidationUtils.verify(true,true,"Currency Amount is verified");
  else
  ValidationUtils.verify(false,true,"Currency Amount is Not Matched ");
  
  if((lowerRange<local_currency_Amount)&&(higherRange>local_currency_Amount))
  ValidationUtils.verify(true,true,"Local Currency Amount is verified");
  else
  ValidationUtils.verify(false,true,"Local Currency Amount is Not Matched ");
  

if((Taxcode1.indexOf("@")!=-1)&&(Taxcode2.indexOf("@")!=-1)){
if(Taxcode1!=""){
var lstIndex = Taxcode1.lastIndexOf("%");
var str = Taxcode1.substring(0, lstIndex);
lstIndex = str.lastIndexOf(" ");
Taxcode1 = str.substring(lstIndex+1).replace(/@/g,'');

var TAC_1 = (parseFloat(currency_Amount)/100)*parseFloat(Taxcode1)
var TAB_1 = (parseFloat(local_currency_Amount)/100)*parseFloat(Taxcode1)

  var lowerRange = parseFloat(TAC_1)-parseFloat("5.00");
  var higherRange = parseFloat(TAC_1)+parseFloat("5.00");

  if(((lowerRange<Tax_Amount_currency_1)&&(higherRange>Tax_Amount_currency_1))||((lowerRange<Tax_Amount)&&(higherRange>Tax_Amount)))
  ValidationUtils.verify(true,true,"Tax Amount Currency 1 is verified");
  else
  ValidationUtils.verify(false,true,"Tax Amount Currency 1 is Not Matched ");
  
}



}
else if(Taxcode1.indexOf("%")!=-1){ 
if(Taxcode1!=""){
var lstIndex = Taxcode1.lastIndexOf("%");
var str = Taxcode1.substring(0, lstIndex);
lstIndex = str.lastIndexOf(" ");
Taxcode1 = str.substring(lstIndex+1).replace(/@/g,'');
//Log.Message(Taxcode1)
var TAC_1 = (parseFloat(currency_Amount)/100)*parseFloat(Taxcode1)
var TAB_1 = (parseFloat(local_currency_Amount)/100)*parseFloat(Taxcode1)

  var lowerRange = parseFloat(TAC_1)-parseFloat("5.00");
  var higherRange = parseFloat(TAC_1)+parseFloat("5.00");
  
//Log.Message(parseFloat(currency_Amount))
//Log.Message(parseFloat(currency_Amount)/100)
//Log.Message(parseFloat(Taxcode1))
//Log.Message(TAC_1)
//Log.Message(Tax_Amount_currency_1)
  if(((lowerRange<Tax_Amount_currency_1)&&(higherRange>Tax_Amount_currency_1))||((lowerRange<Tax_Amount)&&(higherRange>Tax_Amount)))
  ValidationUtils.verify(true,true,"Tax Amount Currency 1 is verified");
  else
  ValidationUtils.verify(false,true,"Tax Amount Currency 1 is Not Matched ");
  
}
}


  
  }
  
  }
  }
  }
if(j<table.getItemCount()-2)
    table.Keys("[Down]");


  }
  
//  var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.Action;
//  action.HoverMouse();
//  ReportUtils.logStep_Screenshot("");
//  action.Click();
////  Delay(3000);
//  action.PopupMenu.Click("Submit Purchase Order");
  
  var SubmitPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
Sys.HighlightObject(SubmitPurchase);
for(var i=0;i<SubmitPurchase.ChildCount;i++){ 
  if((SubmitPurchase.Child(i).isVisible())&&(SubmitPurchase.Child(i).toolTipText=="Submit Purchase Order")){
    SubmitPurchase = SubmitPurchase.Child(i);
    break;
  }
}
WorkspaceUtils.waitForObj(SubmitPurchase);
SubmitPurchase.HoverMouse();
ReportUtils.logStep_Screenshot(); 
SubmitPurchase.Click();

  ReportUtils.logStep_Screenshot();
  aqUtils.Delay(8000, "Submit Purchase Order");;
  TextUtils.writeLog("Submit Purchase Order is Clicked");
  
//var submittedBy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.McTextWidget;
/*
  Sys.Process("Maconomy").Refresh();
  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
  Sys.HighlightObject(table);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  ReportUtils.logStep_Screenshot();
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  Delay(5000);
  */
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  ValidationUtils.verify(true,true,"Purchase Order is Created and Submitted");
//  var PurchaseNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "")
//  .SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).getText();
  ValidationUtils.verify(true,true,"Created Purchase Order Number :"+PONumber);
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("PO Number",EnvParams.Opco,"Data Management",PONumber)

  
}  
}