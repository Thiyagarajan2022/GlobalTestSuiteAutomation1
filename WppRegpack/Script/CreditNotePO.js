//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT EventHandler
//USEUNIT CreateVendorInvoice



/** 
 * This script created Credit note for negative PO
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :03/23/2021
 */
 
Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreditNoteWithPO";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var VendorID,company,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
var Language = "";
var description="";
var OHSN,IHSN,wCodeID,Desp,Qly,UnitPrice ="";
var POnumber = "";
var InvoiceNo,EDate,IDate,Description,TDSValue="";
var Project_manager = "";
var Approved_POnumber = "";
var CreatePO = false;
function CreateNotePO(){ 
TextUtils.writeLog("Credit Note PO is Started"); 
Indicator.PushText("waiting for window to open");

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

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
sheetName = "CreditNoteWithPO";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
Approved_POnumber = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 


ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
if((POnumber==null)||(POnumber=="")){ 
TextUtils.writeLog("Creating Negative Purchase Order for Credit Note");
CreatePO = true;
Runner.CallMethod("CreatePurchaseOrder.CreatePurchaseOrder");
}
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
Log.Message(POnumber)
CreatePO = false;
ExcelUtils.setExcelName(workBook, "Data Management", true);
Approved_POnumber = ReadExcelSheet("Approved Credit PO Number",EnvParams.Opco,"Data Management");
if((Approved_POnumber==null)||(Approved_POnumber=="")){ 
TextUtils.writeLog("Approving Negative Purchase Order for Credit Note");
CreatePO = true;
Runner.CallMethod("ApprovePurchaseOrder.ApprovePurchaseOrder");
}

CreatePO = false;


//try{
//getDetails();
//ExcelUtils.setExcelName(workBook, "Data Management", true);
//  gotoMenu();
//  goToCreatePurchase();
//  gettingApproval();
//  WorkspaceUtils.closeAllWorkspaces();
//    for(var i=level;i<ApproveInfo.length;i++){
//      level=i;
//    WorkspaceUtils.closeMaconomy();
//    aqUtils.Delay(10000, Indicator.Text);
//    var temp = ApproveInfo[i].split("*");
//    Restart.login(temp[2]);
//    aqUtils.Delay(5000, Indicator.Text);
//    todo(temp[3]);
//    FinalApprovePO(temp[1],temp[2],i,temp[3]);
//    }
//POnumber = "1284100108";
   vendorinvoice();
   InvoiceCreation();
   goToAPMenuItem(); 
   invoiceAllocation();
//}
//  catch(err){
//    Log.Message(err);
//  }
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


function getDetails(){ 
  
ExcelUtils.setExcelName(workBook, sheetName, true);
wCodeID = ExcelUtils.getRowDatas("WorkCode",EnvParams.Opco)
Log.Message(wCodeID)
if((wCodeID==null)||(wCodeID=="")){ 
ValidationUtils.verify(false,true,"WorkCode is Needed to Create a Negative Purchase Order");
}

Desp = ExcelUtils.getRowDatas("Description",EnvParams.Opco)
Log.Message(Desp)
if((Desp==null)||(Desp=="")){ 
ValidationUtils.verify(false,true,"Description is Needed to Create a Negative Purchase Order");
}

Qly = ExcelUtils.getRowDatas("Quantity",EnvParams.Opco)
Log.Message(Qly)
if((Qly==null)||(Qly=="")){ 
ValidationUtils.verify(false,true,"Quantity is Needed to Create a Negative Purchase Order");
}
UnitPrice = ExcelUtils.getRowDatas("Cost",EnvParams.Opco)
Log.Message(UnitPrice)
if((UnitPrice==null)||(UnitPrice=="")){ 
ValidationUtils.verify(false,true,"Cost is Needed to Create a Negative Purchase Order");
}
OHSN = ExcelUtils.getRowDatas("Outward HSN",EnvParams.Opco)
//Log.Message(OHSN)
//if((OHSN==null)||(OHSN=="")){ 
//ValidationUtils.verify(false,true,"Outward HSN is Needed to Create a Negative Purchase Order");
//}
IHSN = ExcelUtils.getRowDatas("Inward HSN",EnvParams.Opco)
//Log.Message(IHSN)
//if((IHSN==null)||(IHSN=="")){ 
//ValidationUtils.verify(false,true,"Inward HSN is Needed to Create a Negative Purchase Order");
//}

POS = ExcelUtils.getRowDatas("POS",EnvParams.Opco)
//Log.Message(POS)
//if((POS==null)||(POS=="")){ 
//ValidationUtils.verify(false,true,"POS is Needed to Create a Negative Purchase Order");
//}
  
ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorID = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
if((VendorID=="")||(VendorID==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorID = ExcelUtils.getColumnDatas("Vendor Number",EnvParams.Opco)
}
if((VendorID==null)||(VendorID=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Negative Purchase Order");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Job_Number = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
if((Job_Number=="")||(Job_Number==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_Number = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
}
if((Job_Number==null)||(Job_Number=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to Create a Negative Purchase Order");
}
 
//ExcelUtils.setExcelName(workBook, sheetName, true);
//NOL = ExcelUtils.getColumnDatas("Number of Lines To ADD",EnvParams.Opco)

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
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }
//var allPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Purchase Orders").OleValue.toString().trim())
//WorkspaceUtils.waitForObj(allPurchase);
//allPurchase.Click();

var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
WorkspaceUtils.waitForObj(closefilter);
closefilter.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var craetePurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite;
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
TextUtils.writeLog("Credit Note Purchase Order is created");
ValidationUtils.verify(true,true,"Credit Note Purchase Order is created")

var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-40);
  aqUtils.Delay(5000, Indicator.Text);
  
var ClientCurrency =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McTextWidget;
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

// for(var i=1;i<=10;i++){
//var OHSN,IHSN,wCodeID,Desp,Qly,UnitPrice ="";
//var IHSN ="";
//
//sheetName = "CreditNoteWithPO";
//ExcelUtils.setExcelName(workBook, sheetName, true);
// wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
// Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
// Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
// UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
// OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
// IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
// 
//if((wCodeID!="")||(wCodeID!=null)){
// jB = false; 
// break;
//}
//else{ 
//sheetName = "CreatePurchaseOrder"; 
//init = 0;
//LTA = 10;
//}
//
// }

Log.Message(sheetName)
Log.Message(init)
Log.Message(LTA)


// for(var i=init;i<=LTA;i++){
//var OHSN,IHSN,wCodeID,Desp,Qly,UnitPrice ="";
//var IHSN ="";
//
//if(!jB){ 
// sheetName = "JobBudgetCreation"; 
//}else{ 
// sheetName = "CreatePurchaseOrder"; 
//}
//
//ExcelUtils.setExcelName(workBook, sheetName, true);
// wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
// Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
// Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
// UnitPrice = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
// OHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
// IHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)
////}
//sheetName = "CreatePurchaseOrder";
//ExcelUtils.setExcelName(workBook, sheetName, true);
//var POS = ExcelUtils.getColumnDatas("POS",EnvParams.Opco)


if((wCodeID!="")&&(wCodeID!=null)&&(wCodeID.indexOf("T")==-1)){
//if(line_i<=LTA){
TextUtils.writeLog("Line item "+line_i+" is adding in PO");
line_i++;
var addBudget = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
WorkspaceUtils.waitForObj(addBudget);
addBudget.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var jobNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
WorkspaceUtils.waitForObj(jobNo);
jobNo.Keys("[Tab][Tab]");

var workcode = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
  workcode.Click();
  WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"WorkCode");
  addedlines = true;
  workcode.Keys("[Tab]");
var detailedDescription = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
   if(Desp!=""){
   detailedDescription.setText(Desp);
   }else{ 
   ValidationUtils.verify(false,true,"Detailed Description Needed to create PurchaseOrder");
     }
   detailedDescription.Keys("[Tab]"); 
var Quantity = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(Qly!=""){
   Quantity.setText(Qly);
   }
   Quantity.Keys("[Tab]");
var Unit_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
   if(UnitPrice!=""){
   Unit_Price.setText(UnitPrice);
     }
     
Log.Message("OHSN :"+OHSN);
Log.Message("IHSN :"+IHSN);
Log.Message("POS :"+POS);
  if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_PurchaseOrder.IND_Specific",Unit_Price,OHSN,IHSN,POS);
   
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
WorkspaceUtils.waitForObj(save);
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }


  RowCount++;
}

//}

if(!addedlines)
ValidationUtils.verify(false,true,"WorkCode is not availble in to Create Budget");
else{
  TextUtils.writeLog("Credit Note Purchase Order lines are Saved");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var SubmitPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
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
  var PurchaseNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();;
  ValidationUtils.verify(true,true,"Created Purchase Order Number :"+PurchaseNumber);
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management",PurchaseNumber)
  TextUtils.writeLog("Created Purchase Order Number :"+PurchaseNumber);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}
}



function gettingApproval(){ 
  
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
if((POnumber=="")||(POnumber==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
POnumber = ExcelUtils.getRowDatas("PO Number",EnvParams.Opco)
}
if((POnumber==null)||(POnumber=="")){ 
ValidationUtils.verify(false,true,"PO Number is Needed to Approve Purchase Order");
} 

        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);

  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

//  var allPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Purchase Orders").OleValue.toString().trim())
//  WorkspaceUtils.waitForObj(allPurchase);
//  allPurchase.Click();
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
        
        var mainApprover = ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
        var substitur = ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+POnumber+"*"+ temp;  
//      approvers = EnvParams.Opco+"*"+POnumber+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
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

//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
    Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
sheetName = "CreditNoteWithPO";
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

function vendorinvoice(){
      var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.Click();
    ExcelUtils.setExcelName(workBook, "SSC Users", true);
    var Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    WorkspaceUtils.closeMaconomy();
    Restart.login(Project_manager);  
    }
}

function goToAPMenuItem(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());
}

} 
//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to AP Transactions from Accounts Payable Menu");
TextUtils.writeLog("Entering into AP Transactions from Accounts Payable Menu");
}

function InvoiceCreation(){
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
if((POnumber=="")||(POnumber==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
POnumber = ExcelUtils.getRowDatas("PO Number",EnvParams.Opco)
}
var sheetName = "CreditNoteWithPO";
ExcelUtils.setExcelName(workBook, sheetName, true);

company = EnvParams.Opco;
//company = ExcelUtils.getColumnDatas("Opco",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Vendor Invoice");
}
InvoiceNo = ExcelUtils.getRowDatas("Invoice No",EnvParams.Opco)
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Invoice No. is Needed to Create a Vendor Invoice");
}
Log.Message(InvoiceNo)
EDate = ExcelUtils.getRowDatas("Entry Date",EnvParams.Opco)
if((EDate==null)||(EDate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Vendor Invoice");
}
Log.Message(EDate)
IDate = ExcelUtils.getRowDatas("Invoice Date",EnvParams.Opco)
if((IDate==null)||(IDate=="")){ 
ValidationUtils.verify(false,true,"Invoice Date is Needed to Create a Vendor Invoice");
}
Log.Message(IDate)
description = ExcelUtils.getRowDatas("Description1",EnvParams.Opco)
if((description==null)||(description=="")){ 
ValidationUtils.verify(false,true,"Description1 is Needed to Create a Vendor Invoice");
}
TDSValue = ExcelUtils.getRowDatas("TDS",EnvParams.Opco)

}



function invoiceAllocation(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var allocation = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
allocation.Click(); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
WorkspaceUtils.waitForObj(closefilter);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if(closefilter.text=="Show Filter List"){
}else{ 
  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
  WorkspaceUtils.waitForObj(closefilter);
  closefilter.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var newInvoice = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl2;
WorkspaceUtils.waitForObj(newInvoice);
ReportUtils.logStep_Screenshot();
newInvoice.Click();
aqUtils.Delay(2000, "Waiting for Action");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(5000, "Waiting for Action");
var Create_Method = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McPopupPickerWidget", "", 2);
Create_Method.Keys(" ");
aqUtils.Delay(5000, "Waiting for Action");
Create_Method.Click();
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "From Purchase Order").OleValue.toString().trim(),"Create Method");
aqUtils.Delay(2000, "Waiting for Action");

var Next = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
Sys.HighlightObject(Next);
Next.Click();
aqUtils.Delay(5000, "Waiting for Action");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var PONumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
if(POnumber!=""){
  Log.Message(POnumber)
PONumber.Click();
WorkspaceUtils.SearchByValue_Emp(PONumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),POnumber,"Purchase Order Number");
  }
  
var companyNo = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(company!=""){
companyNo.Click();
WorkspaceUtils.SearchByValue(companyNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),company,"Company Number");
  }
  
var EntryDate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2);
if(EDate!=""){
EntryDate.setText(EDate);
  }
var invoiceDate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2);
if(IDate!=""){
invoiceDate.setText(IDate);
  }
  
  
if(EnvParams.Country.toUpperCase()=="INDIA"){    
var TransactionType = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McValuePickerWidget", "", 2)
if((TransactionType.getText()=="")||(TransactionType.getText()==null)){
TransactionType.Click();
SearchByValue(TransactionType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transaction Type").OleValue.toString().trim(),"Transaction Type");
  }
}

aqUtils.Delay(5000, "Waiting for Action");
var InvoiceType = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McPopupPickerWidget", "", 2);
InvoiceType.Keys(" ");
aqUtils.Delay(5000, "Waiting for Action");
InvoiceType.Click();
Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim())
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim(),"Create Method");
aqUtils.Delay(2000, "Waiting for Action");

//InvoiceType.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim());
aqUtils.Delay(5000, "Waiting for Action");

var InvoiceNumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2)
if(InvoiceNo!=""){
InvoiceNumber.setText(InvoiceNo);
ValidationUtils.verify(true,true,"Invoice No Entered in Invoice Allocation");
   }
   
var Descrip = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
if(description!=""){
Descrip.setText(description);
ValidationUtils.verify(true,true,"Description Entered in Invoice Allocation");
   }
   
var Create = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
Sys.HighlightObject(Create);
Create.Click();

/*
TextUtils.writeLog("New Invoice Button is Clicked");
var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
if(company!=""){
companyNo.Click();
WorkspaceUtils.SearchByValue(companyNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),company,"Company Number");
  }

if(EnvParams.Country.toUpperCase()=="INDIA"){    
var TransactionType = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget3.Composite.McValuePickerWidget;
if((TransactionType.getText()=="")||(TransactionType.getText()==null)){
TransactionType.Click();
SearchByValue(TransactionType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transaction Type").OleValue.toString().trim(),"Transaction Type");
  }
}


var PONumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McValuePickerWidget;
if(POnumber!=""){
  Log.Message(POnumber)
PONumber.Click();
WorkspaceUtils.SearchByValue(PONumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),POnumber,"Purchase Order Number");
  }
var InvoiceType = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite6.McPopupPickerWidget;
//InvoiceType.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice").OleValue.toString().trim());
InvoiceType.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim());
aqUtils.Delay(4000,"Playback")

var EntryDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite3.McDatePickerWidget;
if(EDate!=""){
EntryDate.setText(EDate);
  }
var invoiceDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite5.McDatePickerWidget;
if(IDate!=""){
invoiceDate.setText(IDate);
  }
  
var InvoiceNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.McTextWidget;
if(InvoiceNo!=""){
InvoiceNumber.setText(InvoiceNo);
ValidationUtils.verify(true,true,"Invoice No Entered in Invoice Allocation");
   }
     
var Descrip = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite2.McTextWidget;
if(description!=""){
Descrip.setText(description);
ValidationUtils.verify(true,true,"Description Entered in Invoice Allocation");
   }
   
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
TextUtils.writeLog("Company Number,Purchase Order Number,Entry Date,Description,Invoice Number is Entered and Saved");
*/

aqUtils.Delay(7000, "Waiting for Invoice Allocation");
  p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Invoice Allocation").OleValue.toString().trim())
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Invoice Allocation").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
var Okay = Aliases.Maconomy.Shell7.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Okay.Click();
}


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
excelName = EnvParams.path;
workBook = Project.Path+excelName;
var npEdit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
npEdit.Click();
npEdit.MouseWheel(-1);
//for(var i=2;i<=10;i++){
//PurchOrderNo = ExcelUtils.getColumnDatas("Purch Order No_"+i,EnvParams.Opco)
//var POnumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McValuePickerWidget;
//if(PurchOrderNo!=""){
//POnumber.Click();
//WorkspaceUtils.SearchByValue(POnumber,"Purchase Order",PurchOrderNo,"Purchase Order Number");

//var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
//action.HoverMouse();
//ReportUtils.logStep_Screenshot();
//action.Click();
//aqUtils.Delay(2000, "Waiting for Action");
//action.Click();
//aqUtils.Delay(2000, "Waiting for Action");
//action.Click();
//action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Get Purchase Order").OleValue.toString().trim());
//ReportUtils.logStep_Screenshot();

// }
//  }
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 
  var dueDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.SWTObject("Composite", "", 5).SWTObject("McDatePickerWidget", "", 2);
  dueDate = dueDate.getText().OleValue.toString().trim();
  if(EnvParams.Country.toUpperCase()=="INDIA"){
    CreateVendorInvoice.Language = Language;
  Runner.CallMethod("IND_VendorInvoice.TDS",TDSValue); 
  
  }
  
var curncy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite7.McTextWidget.getText().toString();
Log.Message(curncy);
var grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var tax = grid.getItem(0).getText_2(12).OleValue.toString();
var tax2 = grid.getItem(0).getText_2(14).OleValue.toString();
var tax3 = grid.getItem(0).getText_2(16).OleValue.toString();
var taxcode1 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.Composite.McValuePickerWidget;
if(tax!=taxcode1.getText()){
taxcode1.Click();
WorkspaceUtils.SearchByValue(taxcode1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "G/L Tax Code").OleValue.toString().trim(),tax,"Tax Code 1");
}
var taxcode2 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.Composite2.McValuePickerWidget;
if(tax2!=""){
if(tax2!=taxcode2.getText()){
taxcode2.Click();
WorkspaceUtils.SearchByValue(taxcode2,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "G/L Tax Code").OleValue.toString().trim(),tax2,"Tax Code 2");
}
}
else{ 
 taxcode2.setText(" ") ;
}

var npEdit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
npEdit.Click();
npEdit.MouseWheel(100);

var reaminder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget2.Composite.McTextWidget;
var remainAmount = reaminder.getText().OleValue.toString();
remainAmount=remainAmount.replace("-","");
var amountIncluTax = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite4.McTextWidget;
if(remainAmount!="0.00"){ 
  amountIncluTax.setText(remainAmount);
}
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl3;
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
TextUtils.writeLog("Tax is Validated");
TextUtils.writeLog("Tax Details is Entered and Saved");

aqUtils.Delay(7000, "Waiting for Invoice Allocation");
  p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Invoice Allocation").OleValue.toString().trim())
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Invoice Allocation").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
var Okay = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

Okay.Click();
}
aqUtils.Delay(200, "Waiting for Action")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
action.Click();
aqUtils.Delay(2000, "Waiting for Action");
action.Click();
action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Attach Vendor Document").OleValue.toString().trim());

  TextUtils.writeLog("Document is Attached for Invoice");
  var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Attaching Document");
  var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
  WorkspaceUtils.waitForObj(action);
  action.Click();
  aqUtils.Delay(2000, "Waiting for Action");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(EnvParams.Country.toUpperCase()=="INDIA")
  Runner.CallMethod("IND_VendorInvoice.InvoiceSubmit",action);
  else{
  action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
  WorkspaceUtils.waitForObj(action);
  action.Click();
  aqUtils.Delay(8000, "Waiting for Action");
  action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit for Approval").OleValue.toString().trim());
  
  }
  ReportUtils.logStep_Screenshot();
  aqUtils.Delay(8000, "Submitted for Approval");;
  TextUtils.writeLog("Invoice is Submitted for Approval");
 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var journalNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget.getText().OleValue.toString().trim();
ValidationUtils.verify(true,true,"Created Vendor Invoice Journal Number :"+journalNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("CreditNote Invoice Journal NO",EnvParams.Opco,"Data Management",journalNumber);
ExcelUtils.WriteExcelSheet("CreditNote Vendor Invoice NO",EnvParams.Opco,"Data Management",InvoiceNo);
TextUtils.writeLog("Created Vendor Invoice Journal Number :"+journalNumber);
}










function SearchByValue(ObjectAddrs,popupName,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
//    aqUtils.Delay(3000, Indicator.Text);;
    
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    WorkspaceUtils.waitForObj(code);
    code.setText("Vendor Credit Memo");
//    aqUtils.Delay(3000, Indicator.Text);;
    code.Keys("[Tab]");
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    WorkspaceUtils.waitForObj(code);
    code.setText(EnvParams.Opco);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Search").OleValue.toString().trim()+" ");
    WorkspaceUtils.waitForObj(serch);
//    Sys.HighlightObject(serch);
//    if(serch.isEnabled())
  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
//    aqUtils.Delay(5000, Indicator.Text);;
var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
WorkspaceUtils.waitForObj(OK);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if((table.getItem(i).getText_2(0).OleValue.toString().trim()=="Vendor Credit Memo")&&(table.getItem(i).getText_2(1).OleValue.toString().trim()==EnvParams.Opco)){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
       WorkspaceUtils.waitForObj(OK);
//  if(OK.isEnabled()){
//  OK.HoverMouse();
ReportUtils.logStep_Screenshot();
  OK.Click();
//  }
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//    OK.HoverMouse();
//ReportUtils.logStep_Screenshot();
//   OK.Click(); 
//  }
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
          WorkspaceUtils.waitForObj(cancel);
//if(cancel.isEnabled()){
//  cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
//  }
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//      cancel.HoverMouse();
//ReportUtils.logStep_Screenshot();
//   cancel.Click(); 
//  }
          WorkspaceUtils.waitForObj(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
      WorkspaceUtils.waitForObj(cancel);
//if(cancel.isEnabled()){
//    cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
//  }
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//      cancel.HoverMouse();
//ReportUtils.logStep_Screenshot();
//   cancel.Click(); 
//  }
//      aqUtils.Delay(1000, Indicator.Text);;
WorkspaceUtils.waitForObj(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}



function FinalApprovePO(PONum,Apvr,lvl){ 

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder;
WorkspaceUtils.waitForObj(table);
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
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
