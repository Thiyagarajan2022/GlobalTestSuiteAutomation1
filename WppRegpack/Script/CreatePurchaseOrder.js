//USEUNIT CreditNotePO
//USEUNIT EnvParams
//USEUNIT EventHandler
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT ReverseCreditNote
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
var sheetName = "CreatePurchaseOrder";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
var Language = "";
var Project_manager = "";


//Main Function
function CreatePurchaseOrder(){ 
TextUtils.writeLog("Create Purchase Order Started"); 
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
sheetName = "CreatePurchaseOrder";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
  
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
try{
getDetails();
ExcelUtils.setExcelName(workBook, "Data Management", true);
  gotoMenu();
  goToCreatePurchase();
}
  catch(err){
    Log.Message(err);
  }
WorkspaceUtils.closeAllWorkspaces();
}


function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
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

//Moving to Purchase Order
function gotoMenu(){ 
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountPayable.Exists()){
ImageRepository.ImageSet.AccountPayable.Click();// GL
}
else if(ImageRepository.ImageSet.AccountPayable2.Exists()){
ImageRepository.ImageSet.AccountPayable2.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());
}

} 

ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable Menu");
TextUtils.writeLog("Entering into Purchase Orders from Accounts Payable Menu");
}


//Create and Submit PO
function goToCreatePurchase(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var allPurchase = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
WorkspaceUtils.waitForObj(allPurchase);
allPurchase.Click();

var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
WorkspaceUtils.waitForObj(closefilter);
closefilter.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var craetePurchase = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite;
Sys.HighlightObject(craetePurchase)
for(var i=0;i<craetePurchase.ChildCount;i++){ 
  if((craetePurchase.Child(i).isVisible())&&(craetePurchase.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Purchase Order (Ctrl+N)").OleValue.toString().trim())){
    craetePurchase = craetePurchase.Child(i);
  Log.Message(craetePurchase.FullName)
    break;
  }
}
WorkspaceUtils.waitForObj(craetePurchase);
craetePurchase.HoverMouse();
ReportUtils.logStep_Screenshot(); 
craetePurchase.Click();
TextUtils.writeLog("Create Purchase Order is Clicked");
if(ImageRepository.ImageSet.Img_search.Exists()){ 
  
}

var company = Aliases.Maconomy.Shell6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
if(EnvParams.Opco!=""){
company.Click();
WorkspaceUtils.SearchByValue(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company");
  }else{ 
  ValidationUtils.verify(false,true,"Company Number is Need to create PurchaseOrder");
}
  
var vendor = Aliases.Maconomy.Shell6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
vendor.Click();
SearchByValues_Col_1_all(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorID,"Vendor Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Vendors").OleValue.toString().trim());
  
 
var jobNo = Aliases.Maconomy.Shell6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
jobNo.Click();
WorkspaceUtils.SearchByValues_all_Col_2(jobNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),Job_Number,"Job Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());

var create = Aliases.Maconomy.Shell6.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim())
create.HoverMouse();
ReportUtils.logStep_Screenshot();
create.Click();
TextUtils.writeLog("New Purchase Order is created");
ValidationUtils.verify(true,true,"New Purchase Order is created")

var screen = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-40);
  aqUtils.Delay(5000, Indicator.Text);
  
var ClientCurrency =  
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McTextWidget;
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
sheetName = "CreatePurchaseOrder"; 
init = 0;
LTA = 10;
}

 }

Log.Message(sheetName)
Log.Message(init)
Log.Message(LTA)


 for(var i=init;i<=LTA;i++){
var OHSN,IHSN,wCodeID,Desp,Qly,UnitPrice ="";
var IHSN ="";

if(!jB){ 
 sheetName = "JobBudgetCreation"; 
}else{ 
 sheetName = "CreatePurchaseOrder"; 
}

//sheetName = "JobBudgetCreation";
//ExcelUtils.setExcelName(workBook, sheetName, true);
// wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
// Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
// Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
// UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
// OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
// IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
// 
//if((wCodeID=="")||(wCodeID==null)){
// jB = false; 
//}

 
//if(!jB){
//sheetName = "CreatePurchaseOrder";
ExcelUtils.setExcelName(workBook, sheetName, true);
 wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
 Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
 Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
 UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
 OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
 IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
//}
sheetName = "CreatePurchaseOrder";
ExcelUtils.setExcelName(workBook, sheetName, true);
var POS = ExcelUtils.getColumnDatas("POS",EnvParams.Opco)


if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
//if(line_i<=LTA){
TextUtils.writeLog("Line item "+line_i+" is adding in PO");
line_i++;
var addBudget = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
WorkspaceUtils.waitForObj(addBudget);
addBudget.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var jobNo = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
WorkspaceUtils.waitForObj(jobNo);
jobNo.Keys("[Tab][Tab]");

var workcode = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
  workcode.Click();
  WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"WorkCode");
  addedlines = true;
  workcode.Keys("[Tab]");
var detailedDescription = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
   if(Desp!=""){
   detailedDescription.setText(Desp);
   }else{ 
   ValidationUtils.verify(false,true,"Detailed Description Needed to create PurchaseOrder");
     }
   detailedDescription.Keys("[Tab]"); 
var Quantity = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(Qly!=""){
   Quantity.setText(Qly);
   }
   Quantity.Keys("[Tab]");
var Unit_Price = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3)
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(UnitPrice!=""){
        if(CreditNotePO.CreatePO)
        Unit_Price.setText("-"+UnitPrice);
        else
        Unit_Price.setText(UnitPrice);
     }
     
Log.Message("OHSN :"+OHSN);
Log.Message("IHSN :"+IHSN);
Log.Message("POS :"+POS);
  if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_PurchaseOrder.IND_Specific",Unit_Price,OHSN,IHSN,POS);
   
var save = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
WorkspaceUtils.waitForObj(save);
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
 /* 
aqUtils.Delay(4000, "Validating Tax");
  var tableGrid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var currency_Amount = tableGrid.getItem(RowCount).getText_2(7).OleValue.toString().trim();
  var local_currency_Amount = tableGrid.getItem(RowCount).getText_2(8).OleValue.toString().trim();
  var Taxcode1 = tableGrid.getItem(RowCount).getText_2(12).OleValue.toString().trim();
  var Taxcode2 = tableGrid.getItem(RowCount).getText_2(13).OleValue.toString().trim();
  var Tax_Amount_currency_1 = tableGrid.getItem(RowCount).getText_2(15).OleValue.toString().trim();
  var Tax_Amount_currency_2 = tableGrid.getItem(RowCount).getText_2(17).OleValue.toString().trim();
  var Tax_Amount_1_base = tableGrid.getItem(RowCount).getText_2(14).OleValue.toString().trim();
  var Tax_Amount_2_base = tableGrid.getItem(RowCount).getText_2(16).OleValue.toString().trim();
  var Tax_Amount = tableGrid.getItem(RowCount).getText_2(18).OleValue.toString().trim();
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
    Log.Message(Qly)
    Log.Message(UnitPrice)
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
  Log.Message(lcA)
  lcA = lcA.toFixed(2);
  Log.Message(lcA)
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
  
  
Log.Message(lowerRange) 
Log.Message(higherRange) 
Log.Message(local_currency_Amount)

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
//Log.Message(Taxcode1)
var TAC_1 = (parseFloat(currency_Amount)/100)*parseFloat(Taxcode1)
var TAB_1 = (parseFloat(local_currency_Amount)/100)*parseFloat(Taxcode1)

  var lowerRange = parseFloat(TAC_1)-parseFloat("5.00");
  var higherRange = parseFloat(TAC_1)+parseFloat("5.00");
// Log.Message(lowerRange) 
// Log.Message(higherRange) 
//Log.Message(parseFloat(currency_Amount))
//Log.Message(parseFloat(currency_Amount)/100)
//Log.Message(parseFloat(Taxcode1))
//Log.Message(TAC_1)
//Log.Message(Tax_Amount_currency_1)
// Log.Message(lowerRange) 
// Log.Message(higherRange) 
// Log.Message(Tax_Amount)
  if(((lowerRange<Tax_Amount_currency_1)&&(higherRange>Tax_Amount_currency_1))||((lowerRange<Tax_Amount)&&(higherRange>Tax_Amount)))
  ValidationUtils.verify(true,true,"Tax Amount Currency 1 is verified");
  else
  ValidationUtils.verify(false,true,"Tax Amount Currency 1 is Not Matched ");
  
//  if(TAB_1.toString()==Tax_Amount_1_base.toString())
//  ValidationUtils.verify(true,true,"Tax Amount 1 Base is verified");
//  else
//  ValidationUtils.verify(false,true,"Tax Amount 1 Base is Not Matched ");
}

//if(Taxcode2!=""){
//var lstIndex = Taxcode2.lastIndexOf("%");
//var str = Taxcode2.substring(0, lstIndex);
//lstIndex = str.lastIndexOf(" ");
//Taxcode2 = str.substring(lstIndex+1).replace(/@/g,'');;
//
//var TAC_2 = (parseFloat(currency_Amount)/100)*parseFloat(Taxcode2).toFixed(2)
//var TAB_2 = (parseFloat(local_currency_Amount)/100)*parseFloat(Taxcode2).toFixed(2)
//Log.Message(TAC_1)
//Log.Message(Tax_Amount_currency_1)
//  if(TAC_2.toString()==Tax_Amount_currency_2.toString())
//  ValidationUtils.verify(true,true,"Tax Amount Currency 2 is verified");
//  else
//  ValidationUtils.verify(false,true,"Tax Amount Currency 2 is Not Matched ");
//  
//  if(TAB_2.toString()==Tax_Amount_2_base.toString())
//  ValidationUtils.verify(true,true,"Tax Amount 2 Base is verified");
//  else
//  ValidationUtils.verify(false,true,"Tax Amount 2 Base is Not Matched ");
//  
//
//}

}
else if(Taxcode1.indexOf("%")!=-1){ 
if(Taxcode1!=""){
var lstIndex = Taxcode1.lastIndexOf("%");
var str = Taxcode1.substring(0, lstIndex);
lstIndex = str.lastIndexOf(" ");
Taxcode1 = str.substring(lstIndex+1).replace(/@/g,'');
Log.Message(Taxcode1)
var TAC_1 = (parseFloat(currency_Amount)/100)*parseFloat(Taxcode1)
var TAB_1 = (parseFloat(local_currency_Amount)/100)*parseFloat(Taxcode1)

  var lowerRange = parseFloat(TAC_1)-parseFloat("5.00");
  var higherRange = parseFloat(TAC_1)+parseFloat("5.00");
  
Log.Message(parseFloat(currency_Amount))
Log.Message(parseFloat(currency_Amount)/100)
Log.Message(parseFloat(Taxcode1))
Log.Message(TAC_1)
Log.Message(Tax_Amount_currency_1)
  if(((lowerRange<Tax_Amount_currency_1)&&(higherRange>Tax_Amount_currency_1))||((lowerRange<Tax_Amount)&&(higherRange>Tax_Amount)))
  ValidationUtils.verify(true,true,"Tax Amount Currency 1 is verified");
  else
  ValidationUtils.verify(false,true,"Tax Amount Currency 1 is Not Matched ");
  
//  if(TAB_1.toString()==Tax_Amount_1_base.toString())
//  ValidationUtils.verify(true,true,"Tax Amount 1 Base is verified");
//  else
//  ValidationUtils.verify(false,true,"Tax Amount 1 Base is Not Matched ");
}

}

*/

  RowCount++;

//}

//aqUtils.Delay(4000, "Tax is Validated and Matched");

}

}

if(!addedlines)
ValidationUtils.verify(false,true,"WorkCode is not availble in to Create Budget");
else{
  TextUtils.writeLog("Purchase Order lines are Saved");
//  var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.Action;
//  WorkspaceUtils.waitForObj(action)
//  action.Click();
//  aqUtils.Delay(1000, Indicator.Text);;
//  action.PopupMenu.Click("Submit Purchase Order");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var SubmitPurchase = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
Sys.HighlightObject(SubmitPurchase);
for(var i=0;i<SubmitPurchase.ChildCount;i++){ 
  if((SubmitPurchase.Child(i).isVisible())&&(SubmitPurchase.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit Purchase Order").OleValue.toString().trim())){
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

  ValidationUtils.verify(true,true,"Purchase Order is Created and Submitted");
  var PurchaseNumber = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).getText().OleValue.toString().trim();;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();;
  ValidationUtils.verify(true,true,"Created Purchase Order Number :"+PurchaseNumber);
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  Log.Message(CreditNotePO.CreatePO)
  Log.Message(ReverseCreditNote.CreatePO)
  if(CreditNotePO.CreatePO){ 
  ExcelUtils.WriteExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management",PurchaseNumber)
  TextUtils.writeLog("Created Credit PO Number :"+PurchaseNumber);
  }
  else if(ReverseCreditNote.CreatePO){ 
  ExcelUtils.WriteExcelSheet("ReverseCredit PO Number",EnvParams.Opco,"Data Management",PurchaseNumber)
  TextUtils.writeLog("Created ReverseCredit PO Number :"+PurchaseNumber);
  }
  else{
  ExcelUtils.WriteExcelSheet("PO Number",EnvParams.Opco,"Data Management",PurchaseNumber)
  TextUtils.writeLog("Created Purchase Order Number :"+PurchaseNumber);
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}


}


