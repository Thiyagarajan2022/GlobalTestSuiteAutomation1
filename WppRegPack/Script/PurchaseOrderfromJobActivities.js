//USEUNIT EnvParams
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

ExcelUtils.setExcelName(workBook, sheetName, true);
VendorID = ExcelUtils.getColumnDatas("Vendor Number",EnvParams.Opco)
  if((VendorID=="")||(VendorID==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  VendorID = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
  }
if((VendorID==null)||(VendorID=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Purchase Order");
}

Job_Number = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  if((Job_Number=="")||(Job_Number==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  Job_Number = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  }
if((Job_Number==null)||(Job_Number=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to Create a Purchase Order");
}


  gotoMenu();
  Delay(5000);
  selectJobs();
  listPurchaseOrder();
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
aqUtils.Delay(3000, Indicator.Text);
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

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");

}


function selectJobs(){ 
  var allJobs = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.AllJobs;
  allJobs.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid;
  var firstcell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid.CompanyNumber;
  var closeFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.CloseFilterList;
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid.Jobno;
  job.Click();
  job.setText(Job_Number);
  aqUtils.Delay(7000, Indicator.Text);
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
  closeFilter.Click();
  aqUtils.Delay(8000, Indicator.Text);
  var jobActivities = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.JobActivities;
  jobActivities.Click();
  var iniatePOfromBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.InitiatePO;
  iniatePOfromBudget.Click();
  ReportUtils.logStep_Screenshot("");
  aqUtils.Delay(5000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.POTable.McGrid;
  
   for(var i=1;i<=10;i++){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
  var ChkWorkCode = false;
  if((wCodeID!="")&&(wCodeID!=null)){
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
  if((wCodeID!="")&&(wCodeID!=null)){
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
  aqUtils.Delay(2000, Indicator.Text);
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
  aqUtils.Delay(3000, Indicator.Text);
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
  aqUtils.Delay(4000,Indicator.Text);
  var createPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.CreatePO;
  createPurchaseOrder.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  createPurchaseOrder.Click();
  aqUtils.Delay(8000,Indicator.Text);
  
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
  aqUtils.Delay(4000,Indicator.Text);
}


function listPurchaseOrder(){ 
  var listPO = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.ListPO;
  listPO.Click();
  aqUtils.Delay(5000,Indicator.Text);
  var myPO = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.MyOpenPO;
  myPO.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid;
  var firstcell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid.companyNo;
  var closeFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  
  var POcolumn = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid.POcolumn;
  POcolumn.Click();
  POcolumn.setText(PONumber);
  aqUtils.Delay(7000, Indicator.Text);
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
    aqUtils.Delay(5000, Indicator.Text);
    var POLine = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
    POLine.Click();
    aqUtils.Delay(5000, Indicator.Text);
    
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
  for(var j=0;j<table.getItemCount();j++){
   for(var i=1;i<=10;i++){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
  var Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
  var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
  var UnitPrice = ExcelUtils.getColumnDatas("UnitPrice_"+i,EnvParams.Opco)
  for(var wcL=0;wcL<workcodeList.length;wcL++){
  if((wCodeID!="")&&(wCodeID!=null)&&(workcodeList[wcL].indexOf(wCodeID)!=-1)){
  var temp = workcodeList[wcL].split("*");
  if(table.getItem(j).getText_2(3).OleValue.toString().trim()==temp[1]){ 
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
Delay(2000);
var jobNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
jobNo.Keys("[Tab][Tab]");

var workcode = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  workcode.Click();
  aqUtils.Delay(1000, Indicator.Text);
//  WorkspaceUtils.SearchByValue(workcode,"Work Code",wCodeID,"WorkCode");
  workcode.Keys("[Tab]");
var detailedDescription = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(Desp!=""){
   detailedDescription.setText(Desp);
   ReportUtils.logStep_Screenshot("");
   }else{ 
   ValidationUtils.verify(false,true,"Detailed Description Needed to create PurchaseOrder");
     }
   detailedDescription.Keys("[Tab]"); 
var Quantity = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
   if(Qly!=""){
   Quantity.setText(Qly);
   ReportUtils.logStep_Screenshot("");
   }
   Quantity.Keys("[Tab]");
var Unit_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
   if(UnitPrice!=""){
   Unit_Price.setText(UnitPrice);
   ReportUtils.logStep_Screenshot("");
     }
//  UnitPrice.Keys("[Tab][Tab][Tab]");
  Delay(2000);
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
save.HoverMouse();
ReportUtils.logStep_Screenshot("");
save.Click();
Delay(5000);

  aqUtils.Delay(3000, Indicator.Text);
  
  
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
//  Sys.Desktop.KeyDown(0x10);
//  Sys.Desktop.KeyDown(0x09);
//  Sys.Desktop.KeyUp(0x09);
//  Sys.Desktop.KeyUp(0x10);
//  Sys.Desktop.KeyDown(0x10);
//  Sys.Desktop.KeyDown(0x09);
//  Sys.Desktop.KeyUp(0x09);
//  Sys.Desktop.KeyUp(0x10);
//  Sys.Desktop.KeyDown(0x10);
//  Sys.Desktop.KeyDown(0x09);
//  Sys.Desktop.KeyUp(0x09);
//  Sys.Desktop.KeyUp(0x10);
 
  
  }
  }
  }
  }
if(j<table.getItemCount()-2)
    table.Keys("[Down]");


  }
  
  var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.Action;
  action.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  action.Click();
  Delay(3000);
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
