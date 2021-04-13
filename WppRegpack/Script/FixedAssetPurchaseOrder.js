﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT EventHandler


/**
 * This script create PO for Internal Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :02/12/2021
*/


//Global Varibales
Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "FixedAssetPurchaseOrder";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
var Language = "";
var comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager = "";
var POnumber = "";
var count = true;


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
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "FixedAssetPurchaseOrder";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
POnumber = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 

try{
getPODetail();
ExcelUtils.setExcelName(workBook, "Data Management", true);

if(Job_Number==""){ 
  //Creating Internal JOB
TextUtils.writeLog("Creating Internal Job");
var FolderID = Log.CreateFolder(EnvParams.Opco+"_Internal Job");
Log.PushLogFolder(FolderID);
getDetails();
goToJobMenuItem(); 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }  
createInternalJob();   

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }  
Submiting_Job();

// Approving Job in Multi-Levels
for(var i=0;i<ApproveInfo.length;i++){
level = i;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);

// Restarting maconomy with Approver Logins
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);

//Refreshing To-Do's List to find Submitted Job
vID_Status = true;
todo(temp[3],temp[1],temp[2],i);

//Approving Created Job is every Levels
ApproveJob(temp[1],temp[2],i,temp[3]);
}
WorkspaceUtils.closeAllWorkspaces(); 
Log.PopLogFolder();
TextUtils.writeLog("     ");

}



//Checking Login to execute Job Creation script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();

ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}

//Create PO
var FolderID = Log.CreateFolder(EnvParams.Opco+"_FixedAssetPO");
Log.PushLogFolder(FolderID);
gotoMenu();
goToCreatePurchase();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog("     ");

////Approve PO
//var FolderID = Log.CreateFolder(EnvParams.Opco+"_Approve FixedAssetPO");
//Log.PushLogFolder(FolderID);
//gotoMenu();
//gettingApproval();
//WorkspaceUtils.closeAllWorkspaces();
////CredentialLogin();
//for(var i=level;i<ApproveInfo.length;i++){
//  level=i;
//WorkspaceUtils.closeMaconomy();
//aqUtils.Delay(10000, Indicator.Text);
//var temp = ApproveInfo[i].split("*");
//Restart.login(temp[2]);
//aqUtils.Delay(5000, Indicator.Text);
//todo(temp[3]);
//FinalApprovePO(temp[1],temp[2],i,temp[3]);
//}
//TextUtils.writeLog("Purchase Orders("+POnumber+") is Approved");
//WorkspaceUtils.closeAllWorkspaces();
//Log.PopLogFolder();
//TextUtils.writeLog("     ");


}
  catch(err){
    Log.Message(err);
  }
}


function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}  


function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_group = "Internal";

Job_Type = ExcelUtils.getColumnDatas("Job_Type",EnvParams.Opco)
if((Job_Type==null)||(Job_Type=="")){ 
ValidationUtils.verify(false,true,"Job Type Number is Needed to Create a Internal Job");
}
department = ExcelUtils.getColumnDatas("Department",EnvParams.Opco)
if((department==null)||(department=="")){ 
ValidationUtils.verify(false,true,"Department Number is Needed to Create a Internal Job");
}
buss_unit = ExcelUtils.getColumnDatas("BusinessUnit",EnvParams.Opco)
if((buss_unit==null)||(buss_unit=="")){ 
ValidationUtils.verify(false,true,"BusinessUnit Number is Needed to Create a Internal Job");
}
TemplateNo = ExcelUtils.getColumnDatas("Template",EnvParams.Opco)
if((TemplateNo==null)||(TemplateNo=="")){ 
ValidationUtils.verify(false,true,"Template Number is Needed to Create a Internal Job");
}
Product = ExcelUtils.getColumnDatas("Product",EnvParams.Opco)
if((Product==null)||(Product=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Create a Internal Job");
}
Job_name= ExcelUtils.getColumnDatas("Job_name",EnvParams.Opco)
if((Job_name==null)||(Job_name=="")){ 
ValidationUtils.verify(false,true,"Job Name is Needed to Create a Internal Job");
}
Dlang= ExcelUtils.getColumnDatas("Language",EnvParams.Opco)

BFC= ExcelUtils.getColumnDatas("Counter Party BFC",EnvParams.Opco)

pTerm= ExcelUtils.getColumnDatas("Payment Terms",EnvParams.Opco)
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
}

function getPODetail(){ 
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
Job_Number = ReadExcelSheet("Internal Job Number",EnvParams.Opco,"Data Management");
if((Job_Number=="")||(Job_Number==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_Number = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
}
//if((Job_Number==null)||(Job_Number=="")){ 
//ValidationUtils.verify(false,true,"Job Number is Needed to Create a Purchase Order");
//}
 
ExcelUtils.setExcelName(workBook, sheetName, true);
NOL = ExcelUtils.getColumnDatas("Number of Lines To ADD",EnvParams.Opco)

}


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
for(var i=1;i<=10;i++){
var OHSN,IHSN,wCodeID,Desp,Qly,UnitPrice ="";
var IHSN ="";

sheetName = "FixedAssetPurchaseOrder";
ExcelUtils.setExcelName(workBook, sheetName, true);
 wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
 Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
 Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
 UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
 OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
 IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
var POS = ExcelUtils.getColumnDatas("POS",EnvParams.Opco)


if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
TextUtils.writeLog("Line item "+line_i+" is adding in PO");
line_i++;

var addBudget = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
//var addBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
WorkspaceUtils.waitForObj(addBudget);
addBudget.Click();

var jobNo = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//var jobNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
WorkspaceUtils.waitForObj(jobNo);
jobNo.Keys("[Tab][Tab]");

var workcode = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//var workcode = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
  workcode.Click();
  WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"WorkCode");
  addedlines = true;
  workcode.Keys("[Tab]");
  
var detailedDescription = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//var detailedDescription = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
   if(Desp!=""){
   detailedDescription.setText(Desp);
   }else{ 
   ValidationUtils.verify(false,true,"Detailed Description Needed to create PurchaseOrder");
     }
   detailedDescription.Keys("[Tab]"); 
   
var Quantity = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//var Quantity = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(Qly!=""){
   Quantity.setText(Qly);
   }
   Quantity.Keys("[Tab]");
   
var Unit_Price = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3)
//var Unit_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(UnitPrice!=""){
   Unit_Price.setText(UnitPrice);
     }
     
Log.Message("OHSN :"+OHSN);
Log.Message("IHSN :"+IHSN);
Log.Message("POS :"+POS);
  if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_PurchaseOrder.IND_Specific",Unit_Price,OHSN,IHSN,POS);
   
var save = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
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

var SubmitPurchase = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
//var SubmitPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
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

  ValidationUtils.verify(true,true,"Fixed Purchase Order is Created and Submitted");
  
  var PurchaseNumber = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).getText().OleValue.toString().trim();;
//  var PurchaseNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();;
  ValidationUtils.verify(true,true,"Created Fixed Purchase Order Number :"+PurchaseNumber);
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Fixed Asset PO Number",EnvParams.Opco,"Data Management",PurchaseNumber)
  TextUtils.writeLog("Created Fixed Purchase Order Number :"+PurchaseNumber);
  POnumber = PurchaseNumber;
}


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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}


// Providing Details for Job Creation
function createInternalJob() {
ReportUtils.logStep("INFO", "Enter Job Details");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  

  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var all_job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(all_job);
all_job.HoverMouse();
ReportUtils.logStep_Screenshot("");
  all_job.Click();

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    
  var newJobBtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  WorkspaceUtils.waitForObj(newJobBtn);
  newJobBtn.HoverMouse();
ReportUtils.logStep_Screenshot("");
  newJobBtn.Click();
TextUtils.writeLog("New Job is clicked");
aqUtils.Delay(3000, "Checking Labels");

var cancelJob = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
WorkspaceUtils.waitForObj(cancelJob)

//----------Entering Company Number-------------
ReportUtils.logStep_Screenshot("");
var companyName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,60000);
if(EnvParams.Opco!=""){
companyName.Click();
WorkspaceUtils.SearchByValue(companyName,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company");
  }else{ 
  ValidationUtils.verify(false,true,"Company Number is Need to create PurchaseOrder");
}
  
//----------Entering Job Group-------------

var job = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
if(Job_group!=""){
job.Click();
job.Click();
aqUtils.Delay(3000, "Waiting for ScrolledComposite");
job.Click();
aqUtils.Delay(3000, "Waiting for ScrolledComposite");
var list = "";
try{
list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
}
catch(e){ 
 job.Click(); 
 aqUtils.Delay(3000, "Waiting for ScrolledComposite");
 list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
}
var Add_Visible2 = true;
while(Add_Visible2){
if(list.isEnabled()){
Add_Visible2 = false;
var dropstatus = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Job_group.toString().trim()){ 
          list.Keys("[Enter]");
          aqUtils.Delay(1000, Indicator.Text);
          dropstatus = true;
          ValidationUtils.verify(true,true,"Job Group is listed and selected in Maconomy");
          break;
        }else{ 
          list.Keys("[Down]");
        }
          
      }else{ 
        list.Keys("[Down]");
      }
    }
    if(!dropstatus)
    ValidationUtils.verify(false,true,"Job Group is not listed in Maconomy");
}
}
}
else{ 
    ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");
  }
  


//----------Entering Job Type-------------
   
  var JobType = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Job_Type!=""){
JobType.Click();
WorkspaceUtils.SearchByValue(JobType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Type").OleValue.toString().trim(),Job_Type,"Job Type");
}else{ 
  ValidationUtils.verify(false,true,"JobType is Needed to Create Job");
}

//----------Entering Department-------------   
    
var Depart = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(department!=""){
Depart.Click();
WorkspaceUtils.SearchByValue(Depart,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Department").OleValue.toString().trim(),department,"Department");
}else{ 
  ValidationUtils.verify(false,true,"Department is Needed to Create Job");
}
 
//----------Entering BusinessUnit-------------   
    
  var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(buss_unit!=""){
BussUnit.Click();
WorkspaceUtils.SearchByValue(BussUnit,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Business Unit").OleValue.toString().trim(),buss_unit,"Business Unit");
}else{ 
  ValidationUtils.verify(false,true,"Business Unit is Needed to Create Job");
}

//----------Entering Template Number-------------    
  var template = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(TemplateNo!=""){
template.Click();
var ExlArray = "Internal";
WorkspaceUtils.Config_with_Maconomy_templateValidation(template,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),TemplateNo,ExlArray,Job_Type,comapany,Job_group,"Template Number");
//WorkspaceUtils.SearchByValue(template,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),TemplateNo,"Template");
}else{ 
  ValidationUtils.verify(false,true,"Template is Needed to Create Job");
}
  
    
//----------Entering Product Number-------------    
    
    
var prdNumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Product!=""){
prdNumber.Click();
WorkspaceUtils.SearchByValuePicker_Col_2(prdNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Product Result").OleValue.toString().trim(),Product,"Product Number");
}else{ 
  ValidationUtils.verify(false,true,"Product Number is Needed to Create Job");
}
    
//----------Entering Job Name-------------   
    
    
var jobName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
jobName.setText(Job_name.toString().trim()+" "+STIME);
if((jobName.getText().OleValue.toString().trim()==null)||(jobName.getText().OleValue.toString().trim()==""))
ValidationUtils.verify(false,true,"Job Name can't able to enter in Maconomy");
else
ValidationUtils.verify(true,true,"Job Name is enter in Maconomy");

  
//----------Entering Project Manager-------------
  
var ProjectManger = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);  
 if(Project_manager!=""){
 if(ProjectManger.getText()!=Project_manager.toString().trim()){
 ProjectManger.Click();
 WorkspaceUtils.SearchByValues_Wiz2_Col_2(ProjectManger,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Project Manager");
 }
 }
 
//----------Clicking Create Button or Cancel Button-------------
var btnCreate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());    
if(btnCreate.isEnabled()){

  Sys.HighlightObject(btnCreate)
  btnCreate.HoverMouse();
ReportUtils.logStep_Screenshot("");
  btnCreate.Click();
TextUtils.writeLog("Internal Job is CREATED");
ValidationUtils.verify(true,true,"Internal Job is CREATED");
ReportUtils.logStep("INFO", Job_name+" "+STIME +" : is Created");
TextUtils.writeLog("Job Name :"+Job_name+" "+STIME);
aqUtils.Delay(5000, "Waiting to Check any pop-up for credit limit");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
Log.Message("Language :"+Language);
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

 aqUtils.Delay(10000, "Reporting in HTML about Notification");  
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

 aqUtils.Delay(10000, "Reporting in HTML about Notification");  
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

}
}
else{
  
Log.Message("Language :"+Language);
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

 aqUtils.Delay(10000, "Reporting in HTML about Notification");  
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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
aqUtils.Delay(10000, "Reporting in HTML about Notification"); 

  var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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


aqUtils.Delay(10000, "Reporting in HTML about Notification"); 

  var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

}
  

}

  
}
else{ 
var cancel = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
Sys.HighlightObject(cancel)
cancel.HoverMouse();
ReportUtils.logStep_Screenshot("");
cancel.Click();
ValidationUtils.verify(true,false,"Job is not Created");
ReportUtils.logStep("ERROR", "Job is not Created");
}
aqUtils.Delay(4000, Indicator.Text);
}


//Validating Created Job is available
//Validating Created Job is available
function Submiting_Job() {
  
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  WorkspaceUtils.waitForObj(table);
  var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  WorkspaceUtils.waitForObj(companyFilter);
  companyFilter.forceFocus();
  companyFilter.setVisible(true);
  companyFilter.ClickM();

  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);

  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);

  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  WorkspaceUtils.waitForObj(job);
  job.Click();

  job.setText(Job_name+" "+STIME);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(2000, "Reading Table Data in Job List");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  //Finding Created Job
  flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(Job_name+" "+STIME)){ 
      JobID = table.getItem(v).getText_2(2).OleValue.toString().trim();
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  TextUtils.writeLog("Created Job is available in system");
  TextUtils.writeLog("Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  
  
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  ReportUtils.logStep_Screenshot("");
  closeFilter.Click();
  aqUtils.Delay(4000, "Created Job Details is loading");
  
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  //Validating Job and Invoice Currency
  aqUtils.Delay(2000, "Checking Job and Invoice Currency");
  var JobCurrency = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  var InvoiceCurrency = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  JobCurrency = JobCurrency.getText().OleValue.toString().trim();
  InvoiceCurrency = InvoiceCurrency.getText().OleValue.toString().trim();
  ReportUtils.logStep("Job Currency:"+JobCurrency)
  ReportUtils.logStep("Invoice Currency:"+InvoiceCurrency)
  Log.Message("JobCurrency :"+JobCurrency);
  Log.Message("InvoiceCurrency :"+InvoiceCurrency);
  
  
  //Changing Invoice Currency as Job Currency
    if(JobCurrency!=InvoiceCurrency){ 
    var prices =  Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 12)
    Sys.HighlightObject(prices);
    prices.Click(); 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
  
  aqUtils.Delay(2000, "Changing Job and Invoice Currency");
  var InvoiceCurr = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 11).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
  WorkspaceUtils.waitForObj(InvoiceCurr);
  InvoiceCurr.Keys(JobCurrency);
  aqUtils.Delay(4000, "Changing Invoice Currency as "+JobCurrency);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var Save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
  Sys.HighlightObject(Save)
  Save.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  }
  
  // Moving to Information Tab to Submit
  var info = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 5)
  Sys.HighlightObject(info);
  info.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  info.Click();
  count=false;
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var counterBFC = Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if((counterBFC.getText()=="") || (counterBFC.getText()==null)){ 
  counterBFC.Click();
  WorkspaceUtils.SearchByValue(counterBFC,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Counter Party BFC").OleValue.toString().trim(),BFC,"Counter Party BFC");

  }
  aqUtils.Delay(3000, "Submitting Job");
  var AccountDirector = Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if((AccountDirector.getText()=="")||(AccountDirector.getText()==null)){ 
  AccountDirector.Click();
  WorkspaceUtils.SearchByValues_Wiz2_Col_2(AccountDirector,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Project Manager");
  }
  aqUtils.Delay(3000, "Submitting Job");
  var Submit = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8);
  Sys.HighlightObject(Submit);
  Submit.Click();
  aqUtils.Delay(4000, "Submitting Job for Approval");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var Sliding_Panel = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
  Sys.HighlightObject(Sliding_Panel);
  Sliding_Panel.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  ImageRepository.ImageSet.Maximize.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var Job_Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
  Job_Approve.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
  
  var Approval_Table = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(Approval_Table);
    var y=0;
    
    //Getting User Name
    Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
  for(var i=0;i<Approval_Table.getItemCount();i++){   
     var approvers="";
      if(Approval_Table.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
        
      //Self Approve is Disabled. So finding different Approver
        var mainApprover = Approval_Table.getItem(i).getText_2(7).OleValue.toString().trim();
        var substitur = Approval_Table.getItem(i).getText_2(8).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+JobID+"*"+ temp;
      Log.Message("Approver level :" +i+ ": " +approvers);
      Approve_Level[y] = approvers;
      y++;
      }
}
TextUtils.writeLog("Finding approvers for Created Job");
CredentialLogin();
  Log.Message(JobID)
  Log.Message(TemplateNo)

}
}

//Refreshing To-Do's List and Seleting Notification of Jobs
function todo(lvl,JobID,Apvr){ 
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
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job from To-Dos List"); 
listPass = false; 

//Finding Job in Approve Job Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
}


if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job (Substitute) from To-Dos List");  
var listPass = false;   
//Finding Job in Approve Job (Substitute) Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
}  



if((listPass)||(!flag)){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job by Type from To-Dos List"); 
listPass = false; 

//Finding Job in Approve Job by Type Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
}


if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job by Type (Substitute) from To-Dos List"); 
var listPass = false;   

//Finding Job in Approve Job by Type (Substitute) Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
} 
  }
  
}

//Finding Job is available in Approvers
function FinalApproveJob(JobID,Apvr,lvl){ 
  

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
//Finding Screen with Close Filter or Show Filter
var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
var showFilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
WorkspaceUtils.waitForObj(showFilter);
showFilter.Click();

}


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(table);
var firstCell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
firstCell.setText(JobID);
var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(5000, Indicator.Text);;
var i=0;
var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);

flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==JobID){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}


}




//Approve the Created Job
function ApproveJob(JobID,Apvr,lvl){ 
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }


if(flag){
vID_Status = false;
ValidationUtils.verify(flag,true,"Created Job is available in Approval List");
TextUtils.writeLog("Created Job is available in Approval List");
var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;

closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var Approve = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)

Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Sys.HighlightObject(Approve)
Approve.Click();


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(2000, "Approving Job");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
ValidationUtils.verify(true,true,"Job is Approved by "+Apvr)
TextUtils.writeLog("Job is Approved by "+Apvr);

// After Final Approve validating in Sliding panel
if(Approve_Level.length==lvl+1){
var approvalBar = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();

  aqUtils.Delay(2000, "Approving Job");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  ImageRepository.ImageSet.Maximize.Click();
    

  aqUtils.Delay(2000, "Approving Job");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var JobApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2

WorkspaceUtils.waitForObj(JobApproval);
JobApproval.HoverMouse();
ReportUtils.logStep_Screenshot();
JobApproval.Click();
aqUtils.Delay(3000, Indicator.Text);;

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

WorkspaceUtils.waitForObj(approvertable);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
ReportUtils.logStep_Screenshot();

for(var i=0;i<approvertable.getItemCount();i++){   
var approvers="";
if(approvertable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(true,false,"Job is not Approved in Level : "+i);
}
}

var JobApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl;
Sys.HighlightObject(JobApproval);
JobApproval.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

ImageRepository.ImageSet.Forward.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var Blanket_Invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Estimating = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Time_Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Purchase_Ordering = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Amount_Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Invoicing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 9).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");

//var Blanket_Invoice = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Estimating = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Time_Registration = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Purchase_Ordering = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Amount_Registration = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Invoicing = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 9).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");

Sys.HighlightObject(Blanket_Invoice)
Sys.HighlightObject(Estimating)
Sys.HighlightObject(Time_Registration)
Sys.HighlightObject(Purchase_Ordering)
Sys.HighlightObject(Amount_Registration)
Sys.HighlightObject(Invoicing)

var checkmark = false;

    if(Blanket_Invoice.getSelection()){
      Blanket_Invoice.Click();
      ValidationUtils.verify(true,true,"Job for Blanket Invoice is Unchecked");
      checkmark = true;
    } 
    
    if(Estimating.getSelection()){
      Estimating.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Estimating is Unchecked");
      checkmark = true;
    } 
    
    if(Time_Registration.getSelection()){
      Time_Registration.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Time Registration is Unchecked");
      checkmark = true;
    } 
    
    if(Purchase_Ordering.getSelection()){
      Purchase_Ordering.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Purchase Ordering is Unchecked");
      checkmark = true;
    } 
    
     if(Amount_Registration.getSelection()){
      Amount_Registration.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Amount Registration is Unchecked");
      checkmark = true;
    } 
    
     if(Invoicing.getSelection()){
      Invoicing.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Invoicing is Unchecked");
      checkmark = true;
    } 
    
    if(checkmark){ 
      aqUtils.Delay(3000, "Saving Changes for Job");;
      var Save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      Sys.HighlightObject(Save);
      Save.Click();
      aqUtils.Delay(3000, "Saving Changes for Job");;
    }
    
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Internal Job Number",EnvParams.Opco,"Data Management",JobID)

//var Save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//      Sys.HighlightObject(Save);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}

  ValidationUtils.verify(true,true,"Job is Approved by "+Apvr)

}



}


//==========Approve PO=========================================

function gettingApproval(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

  var allPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
  WorkspaceUtils.waitForObj(allPurchase);
  allPurchase.Click();
  var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.CompanyNo;
  WorkspaceUtils.waitForObj(companyNo);
  companyNo.Click();
  companyNo.Keys("[Tab]");
  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.PurchaseNumber;
  WorkspaceUtils.waitForObj(purchaseNo);
  purchaseNo.Click();
  purchaseNo.setText(POnumber);

  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid;
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(3000, "Reading Table Data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==POnumber){ 
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
  ValidationUtils.verify(flag,true,"Created Purchase Order is available in system");
  TextUtils.writeLog("Created Purchase Order is available in system");
  
  
 if(flag){
  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(closefilter);
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  WorkspaceUtils.waitForObj(purchaseOrderApproval);
  purchaseOrderApproval.Click();
  
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }

  ImageRepository.ImageSet.Maximize.Click();

  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
  WorkspaceUtils.waitForObj(purchaseApproval);
  purchaseApproval.Click();
  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(ApproverTable);
  var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      approvers = EnvParams.Opco+"*"+POnumber+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
      Approve_Level[y] = approvers;
      y++;
      }
}
TextUtils.writeLog("Finding approvers for Created Purchase Order");
var listPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.listPurchaseOrder.TabControl;
listPurchaseOrder.Click();
ImageRepository.ImageSet.Forward.Click();


CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");

ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
//sheetName = "ApprovePurchaseOrder";
if(OpCo2[2]==Project_manager){
level = 1;
var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
}
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
//  Approve.PopupMenu.Click("Approve Purchase Order");
//ImageRepository.ImageSet.ApprovePurchaseOrder.Click();
//ReportUtils.logStep_Screenshot();
//aqUtils.Delay(8000, Indicator.Text);;
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created PO");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);

  ValidationUtils.verify(true,true,"Purchase Order is Approved by :"+loginPer)
  TextUtils.writeLog("Purchase Order is Approved by :"+loginPer); 

  
if(Approve_Level.length==1){
  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  purchaseOrderApproval.Click();

  ImageRepository.ImageSet.Maximize.Click();

  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
  WorkspaceUtils.waitForObj(purchaseApproval)
  purchaseApproval.Click();
  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable)
ReportUtils.logStep_Screenshot();
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


function todo_PO(lvl){ 
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
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Purchase Order (Substitute) from To-Dos List");  
var listPass = false;   
  }
}  


if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Purchase Order by Type from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Purchase Order by Type (Substitute) from To-Dos List"); 
var listPass = false;   
  }
} 
  }
  
}

function FinalApprovePO(PONum,Apvr,lvl){ 

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder;
WorkspaceUtils.waitForObj(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid;
var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox;
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(PONum);
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter;
WorkspaceUtils.waitForObj(table);
aqUtils.Delay(3000, "Reading Data from table");;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==PONum){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Purchase Order is available in Approval List");
TextUtils.writeLog("Created Purchase Order is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2;
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
}
WorkspaceUtils.waitForObj(Approve);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
//aqUtils.Delay(3000, Indicator.Text);
//Approve.PopupMenu.Click("Approve Purchase Order");
//ReportUtils.logStep_Screenshot();
//aqUtils.Delay(8000, Indicator.Text);;
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Apvr)
TextUtils.writeLog("Purchase Order is Approved by "+Apvr);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//Uncommand
/*
//var screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");;
var screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "")
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-5);
//  aqUtils.Delay(5000, Indicator.Text);
//var ApvPerson = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 6).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
var ApvPerson = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 6).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while ((ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

    if(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1){
  ValidationUtils.verify(true,true,"Purchase Order is Approved by :"+loginPer)
  TextUtils.writeLog("Purchase Order is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Purchase Order is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Purchase Order is Approved by :"+loginPer+ "But its Not Reflected")
  }
  
*/


if(Approve_Level.length==lvl+1){
var approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.Approvals;
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
ImageRepository.ImageSet.Maximize.Click();

if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN"))
   Runner.CallMethod("IND_ApprovePurchaseOrder.ApprovalStatus");
else{
var POapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
WorkspaceUtils.waitForObj(POapproval)
POapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
POapproval.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
WorkspaceUtils.waitForObj(approvertable)
ReportUtils.logStep_Screenshot();
}
}
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Apvr)
}
}

}


