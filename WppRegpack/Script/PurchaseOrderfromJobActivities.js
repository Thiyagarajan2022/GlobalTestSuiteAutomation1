﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/**
 * This script create PO for Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :02/12/2021
*/

//Global Varibales
Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "POFromJobBudget";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
var PONumber = "";
var workcodeList = [];
var Language = "";
var Project_manager = ""


//Main Function
function CreatePurchaseOrder(){ 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;


//Checking Login to execute Job Creation script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "POFromJobBudget";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
PONumber = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);

//try{
getDetails();
ExcelUtils.setExcelName(workBook, sheetName, true);
  gotoMenu();
  selectJobs();
  listPurchaseOrder();
//}
//  catch(err){
//    Log.Message(err);
//  }
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
closeAllWorkspaces();
}


//Getting Details from Datasheet
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


ExcelUtils.setExcelName(workBook, sheetName, true);
NOL = ExcelUtils.getColumnDatas("Number of Lines To ADD",EnvParams.Opco)

}


//Moving to Purchase Order from WorkSpace Client
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Jobs").OleValue.toString().trim());
}

} 


ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Moved to Jobs from Jobs Menu");
}



//Selecting Job
function selectJobs(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var allJobs = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  allJobs.Click();
  ReportUtils.logStep_Screenshot("");
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
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//  aqUtils.Delay(3000, Indicator.Text);
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
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, Indicator.Text);
  var jobActivities = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.JobActivities;
  WorkspaceUtils.waitForObj(jobActivities);
  jobActivities.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, Indicator.Text);
  var iniatePOfromBudget = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 11)
//  var iniatePOfromBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.SWTObject("TabControl", "", 11);
  WorkspaceUtils.waitForObj(iniatePOfromBudget);
  iniatePOfromBudget.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, Indicator.Text);
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Initiate Purchase Order is clicked");
  var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("Composite", "", 10).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.SWTObject("Composite", "", 10).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
  WorkspaceUtils.waitForObj(table);

var RowCount = 0;
var addedlines = false;
var jB = true;
var line_i =1;
var LTA = 1;
var init = 1;
if((NOL==null)||(NOL==""))
{ 
  LTA = 10;
}else{ 
  if(NOL.indexOf("-")!=-1){ 
    var line_Temp = NOL.split("-");
    init = aqConvert.StrToInt(line_Temp[0]);
    LTA = aqConvert.StrToInt(line_Temp[1]);
  }else{
  LTA = aqConvert.StrToInt(NOL);
  }
}

Log.Message(init)
Log.Message(LTA)

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

 for(var i=1;i<=10;i++){
var OHSN,IHSN,wCodeID,Desp,Qly,UnitPrice ="";
var IHSN ="";

sheetName = "JobBudgetCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
 wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
 Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
 Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
 UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
 OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
 IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
 
if((wCodeID!="")||(wCodeID!=null)){
 jB = false; 
 break;
}
else{ 
sheetName = "POFromJobBudget"; 
init = 0;
LTA = 10;
}

 }

Log.Message(sheetName)
Log.Message(init)
Log.Message(LTA)


  workcodeList = [];
  var wcL = 0;
  for(var j=0;j<table.getItemCount();j++){
       for(var i=init;i<=LTA;i++){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
  if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
  if(table.getItem(j).getText_2(3).OleValue.toString().trim()==wCodeID){ 
    if(line_i<=LTA){
    line_i++;
  workcodeList[wcL] = table.getItem(j).getText_2(3).OleValue.toString().trim()+"*"+table.getItem(j).getText_2(4).OleValue.toString().trim();
  wcL++;
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  var selected = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("Composite", "", 10).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
//  var selected = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.SWTObject("Composite", "", 10).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  if(!selected.getSelection()){ 
  selected.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  selected.Click();
  }
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  // 8 TAB to move to Vendor
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
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(1000, Indicator.Text);
  var vendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("Composite", "", 10).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4)
//  var vendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.SWTObject("Composite", "", 10).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
  vendor.Click();
  WorkspaceUtils.SearchByValues_Col_1_all(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Vendor").OleValue.toString().trim(),VendorID,"Vendor Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "All Vendors").OleValue.toString().trim());
  aqUtils.Delay(5000, Indicator.Text);
  var save = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.newassetbutton
//  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.RemarksSave;
  save.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  save.Click();
  aqUtils.Delay(3000, "Saving Changes");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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
  }// to check Number of lines
  
  }
  }
  }
if(j<table.getItemCount()-2)
    table.Keys("[Down]");


  }
  
  }
//  aqUtils.Delay(4000,Indicator.Text);
  var createPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SWTObject("SingleToolItemControl", "", 12)
//  var createPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.CreatePO;
  WorkspaceUtils.waitForObj(createPurchaseOrder);
  createPurchaseOrder.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  createPurchaseOrder.Click();
//  aqUtils.Delay(8000,Indicator.Text);
  
  var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Jobs").OleValue.toString().trim()).SWTObject("Label", "*");
  ReportUtils.logStep(label.getText());
  Log.Message(label.getText())
  PONumber = label.WndCaption;
  PONumber = PONumber.substring(PONumber.lastIndexOf(" ")+1);
  ValidationUtils.verify(true,true,"Purchase Order Number :"+PONumber);   
  var Okay = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Jobs").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "OK").OleValue.toString().trim());
  Okay.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  Okay.Click();
  aqUtils.Delay(2000,"Created PO Number :"+PONumber);
  TextUtils.writeLog("Purchase Order is Created :"+PONumber);
}


function listPurchaseOrder(){ 
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var listPO = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl2
//  var listPO = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.ListPO;
  WorkspaceUtils.waitForObj(listPO);
  listPO.Click();
//  aqUtils.Delay(5000,Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000,"Finding Created PO Number :"+PONumber);
  var myPO = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "My POs").OleValue.toString().trim());
//  var myPO = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "My POs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(myPO);
  myPO.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
//  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid;
  var firstcell = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
//  var firstcell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid.companyNo;
  var closeFilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
//  var closeFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(6000, Indicator.Text);
  var POcolumn = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
//  var POcolumn = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.POTable.McGrid.POcolumn;
  POcolumn.Click();
  POcolumn.setText(PONumber);
  WorkspaceUtils.waitForObj(table);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var screen = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
//  var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-5);
  aqUtils.Delay(3000, "Reading Vendor Currency");
//var ClientCurrency =  Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2)
//var ClientCurrency =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McTextWidget;
//WorkspaceUtils.waitForObj(ClientCurrency);

ExcelUtils.setExcelName(workBook, "Data Management", true);
ClientCurrency = ReadExcelSheet("Global Vendor Currency",EnvParams.Opco,"Data Management");
//ClientCurrency =  ClientCurrency.getText();  
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
    
    
    var POLine = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl
//    var POLine = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
    WorkspaceUtils.waitForObj(POLine);
    POLine.Click();
//    aqUtils.Delay(5000, Indicator.Text);  

  var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid  
//  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
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
  Log.Message(workBook)
  Log.Message(sheetName) 

  ExcelUtils.setExcelName(workBook, "POFromJobBudget", true);
  var POS = ExcelUtils.getColumnDatas("POS",EnvParams.Opco)


  Log.Message(wCodeID)
  Log.Message(Desp)
  Log.Message(Qly)
  Log.Message(UnitPrice)
  Log.Message(OHSN)
  Log.Message(IHSN)
  Log.Message(POS)
  
  for(var wcL=0;wcL<workcodeList.length;wcL++){
  if((wCodeID!="")&&(wCodeID!=null)&&(workcodeList[wcL].indexOf(wCodeID)!=-1)){
  var temp = workcodeList[wcL].split("*");
  if(table.getItem(j).getText_2(3).OleValue.toString().trim()==temp[1]){ 
    Log.Message(temp[1])
    aqUtils.Delay(3000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(3000, Indicator.Text);
//Delay(2000);
var jobNo = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//var jobNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
jobNo.Keys("[Tab][Tab]");
aqUtils.Delay(3000, Indicator.Text);

var workcode = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//var workcode = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  workcode.Click();
  aqUtils.Delay(1000, Indicator.Text);
//  WorkspaceUtils.SearchByValue(workcode,"Work Code",wCodeID,"WorkCode");
  workcode.Keys("[Tab]");
  
var detailedDescription = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid.Remarks
//var detailedDescription = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;

//Log.Message(Desp)
//Log.Message(detailedDescription.getText())
//Log.Message(Desp!="")
//Log.Message(detailedDescription.getText()!=Desp)
//Log.Message((Desp!="")&&(detailedDescription.getText()!=Desp))
   if((Desp!="")&&(detailedDescription.getText()!=Desp)){
   detailedDescription.setText(Desp);
   ReportUtils.logStep_Screenshot("");
   }
   detailedDescription.Keys("[Tab]"); 
var Quantity = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid.Remarks
//var Quantity = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
   if((Qly!="")&&(Quantity.getText()!=Qly)){
   Quantity.setText(Qly);
   ReportUtils.logStep_Screenshot("");
   }
   Quantity.Keys("[Tab]");
   
var Unit_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid.Remarks
//var Unit_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
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

var save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
//var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  
if(EnvParams.Country.toUpperCase()=="INDIA"){
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  Sys.Desktop.KeyDown(0x10);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x10);
  aqUtils.Delay(1000, "Saving Changes");
  }
 

aqUtils.Delay(4000, "Validating Tax");
/*
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
    lcA = parseFloat(Qly)*parseFloat(UnitPrice);
    Log.Message(lcA)
  }
  else if(ClientCurrency!="GBP"){
   convertCurr = 1/BaseCurrency;
  var QtyXCurr = parseFloat(convertCurr)*parseFloat(CA);
   lcA = parseFloat(QtyXCurr)*parseFloat(ExchangeRate);
  }
  else{ 
    lcA = parseFloat(CA)*parseFloat(ExchangeRate);
  }

  lcA = lcA.toFixed(2);
  var lowerRange = parseFloat(lcA)-parseFloat("1000.00");
  var higherRange = parseFloat(lcA)+parseFloat("1000.00");

  if((Taxcode1=="")&&(Taxcode2==""))
  ValidationUtils.verify(false,true,"Tax Code 1 and Tax Code 2 is not Populated");
  if(Taxcode1!="")
  ValidationUtils.verify(true,true,"Tax Code 1 is populated");
  if(Taxcode2!="")
  ValidationUtils.verify(true,true,"Tax Code 2 is populated");
  
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

*/

  
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
  
  var screen = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10
//  var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(40);
  aqUtils.Delay(3000, "Validating Purchaser");
var purchaserName = NameMapping.Sys.Maconomy.ChangeEmployee.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
//var purchaserName = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
 if(purchaserName.getText()!=Project_manager.toString().trim()){
 purchaserName.Click();
 WorkspaceUtils.SearchByValues_Wiz2_Col_2(purchaserName,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Purchaser");
 }

  var SubmitPurchase = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
//  var SubmitPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
Sys.HighlightObject(SubmitPurchase);
for(var i=0;i<SubmitPurchase.ChildCount;i++){ 
  if((SubmitPurchase.Child(i).isVisible())&&(SubmitPurchase.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Submit Purchase Order").OleValue.toString().trim())){
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
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  
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
