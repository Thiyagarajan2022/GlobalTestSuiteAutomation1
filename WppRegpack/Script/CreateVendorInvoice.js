﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
 

/** 
 * This script create Vendor Invoice
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :03/23/2021
 */
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "VendorInvoice";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var company,PurchOrderNo,InvoiceNo,Description,IDate,EDate,TDSValue ="";
var Language = "";
var paymentMode = "";
var ExlAmount = "";
var remainAmount = "";

function CreateInvoice(){ 
TextUtils.writeLog("Create Vendor Invoice Started"); 
Indicator.PushText("waiting for window to open");

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "VendorInvoice";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
company,PurchOrderNo,InvoiceNo,Description,IDate,EDate,TDSValue ="";
paymentMode = "";
ExlAmount = "";
remainAmount = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
//try{
getDetails();
goToJobMenuItem(); 
invoiceAllocation();
//}catch(err){ 
//  Log.Message(err);
//}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

function goToJobMenuItem(){ 
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


function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);

company = ExcelUtils.getColumnDatas("Opco",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Vendor Invoice");
}
//Log.Message(company)
ExcelUtils.setExcelName(workBook, "Data Management", true);
PurchOrderNo = ReadExcelSheet("PO Number",EnvParams.Opco,"Data Management");
if((PurchOrderNo=="")||(PurchOrderNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
PurchOrderNo = ExcelUtils.getColumnDatas("Purch Order No",EnvParams.Opco)
}
if((PurchOrderNo==null)||(PurchOrderNo=="")){ 
ValidationUtils.verify(false,true,"PO Number is Needed to Create a Vendor Invoice");
}
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(PurchOrderNo)
InvoiceNo = ExcelUtils.getColumnDatas("Invoice No",EnvParams.Opco)
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Invoice No. is Needed to Create a Vendor Invoice");
}
//Log.Message(InvoiceNo)
EDate = ExcelUtils.getColumnDatas("Entry Date",EnvParams.Opco)
if((EDate==null)||(EDate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Vendor Invoice");
}
//Log.Message(EDate)
IDate = ExcelUtils.getColumnDatas("Invoice Date",EnvParams.Opco)
if((IDate==null)||(IDate=="")){ 
ValidationUtils.verify(false,true,"Invoice Date is Needed to Create a Vendor Invoice");
}
//Log.Message(IDate)
Description = ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
if((Description==null)||(Description=="")){ 
ValidationUtils.verify(false,true,"Description is Needed to Create a Vendor Invoice");
}
TDSValue = ExcelUtils.getColumnDatas("TDS",EnvParams.Opco)
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
var POnumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
if(PurchOrderNo!=""){
POnumber.Click();
WorkspaceUtils.SearchByValue_Emp(POnumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),PurchOrderNo,"Purchase Order Number");
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
InvoiceType.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice").OleValue.toString().trim());
aqUtils.Delay(5000, "Waiting for Action");

var InvoiceNumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2)
if(InvoiceNo!=""){
InvoiceNumber.setText(InvoiceNo);
ValidationUtils.verify(true,true,"Invoice No Entered in Invoice Allocation");
   }
   
var Descrip = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
if(Description!=""){
Descrip.setText(Description);
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


var POnumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McValuePickerWidget;
if(PurchOrderNo!=""){
POnumber.Click();
WorkspaceUtils.SearchByValue(POnumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),PurchOrderNo,"Purchase Order Number");
  }
var InvoiceType = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite6.McPopupPickerWidget;
InvoiceType.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice").OleValue.toString().trim());

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
if(Description!=""){
Descrip.setText(Description);
ValidationUtils.verify(true,true,"Description Entered in Invoice Allocation");
   }
   
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();

*/

TextUtils.writeLog("Company Number,Purchase Order Number,Entry Date,Description,Invoice Number is Entered and Saved");
//if(ImageRepository.ImageSet.OK_Button.Exists()){ 
//var Okay = Aliases.Maconomy.Shell7.Composite.Button;
//Okay.Click();
//}

Delay(5000)
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Invoice Allocation").OleValue.toString().trim())
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Invoice Allocation").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
var Okay = Aliases.Maconomy.Shell7.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim()
)
Okay.Click();
}


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
excelName = EnvParams.path;
workBook = Project.Path+excelName;
var npEdit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
npEdit.Click();
npEdit.MouseWheel(-1);
for(var i=2;i<=10;i++){
PurchOrderNo = ExcelUtils.getColumnDatas("Purch Order No_"+i,EnvParams.Opco)
var POnumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.McValuePickerWidget;
if(PurchOrderNo!=""){
POnumber.Click();
WorkspaceUtils.SearchByValue(POnumber,"Purchase Order",PurchOrderNo,"Purchase Order Number");

var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
action.HoverMouse();
ReportUtils.logStep_Screenshot();
action.Click();
aqUtils.Delay(2000, "Waiting for Action");
action.Click();
aqUtils.Delay(2000, "Waiting for Action");
action.Click();
//aqUtils.Delay(3000, Indicator.Text);
action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Get Purchase Order").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();

////  Sys.Process("Maconomy").Refresh();
//  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
//  Sys.HighlightObject(table);
//  Sys.Desktop.KeyDown(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyUp(0x28);
//  Delay(1000);
//  ReportUtils.logStep_Screenshot();
//  Sys.Desktop.KeyDown(0x0D);
//  Sys.Desktop.KeyUp(0x0D);

//aqUtils.Delay(5000, Indicator.Text);
  }
  }
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
 
  var dueDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.SWTObject("Composite", "", 5).SWTObject("McDatePickerWidget", "", 2);
  dueDate = dueDate.getText().OleValue.toString().trim();
  if(EnvParams.Country.toUpperCase()=="INDIA"){
  Runner.CallMethod("IND_VendorInvoice.TDS",TDSValue); 
  
  }
  
var curncy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite7.McTextWidget.getText().toString();
Log.Message(curncy);
var grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var tax = grid.getItem(0).getText_2(12).OleValue.toString();
var tax2 = grid.getItem(0).getText_2(14).OleValue.toString();
var tax3 = grid.getItem(0).getText_2(16).OleValue.toString();
var taxcode1 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.Composite.McValuePickerWidget;
if(tax!=""){
if(tax!=taxcode1.getText()){
taxcode1.Click();
WorkspaceUtils.SearchByValue(taxcode1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "G/L Tax Code").OleValue.toString().trim(),tax,"Tax Code 1");
}
}
else{ 
taxcode1.Click();
taxcode1.setText(" ");
}

var taxcode2 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.Composite2.McValuePickerWidget;
if(tax2!=""){
if(tax2!=taxcode2.getText()){
taxcode2.Click();
WorkspaceUtils.SearchByValue(taxcode2,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "G/L Tax Code").OleValue.toString().trim(),tax2,"Tax Code 2");
}
}
else{ 
taxcode2.Click();
taxcode2.setText(" ");
}

var Save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl3;
WorkspaceUtils.waitForObj(Save);
Save.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//var taxcode3 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.Composite3.McValuePickerWidget;
//if(tax3!=taxcode3.getText()){
//taxcode3.Click();
//WorkspaceUtils.SearchByValue(taxcode3,"G/L Tax Code",tax3,"Tax Code 3");
//}
var npEdit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
npEdit.Click();
npEdit.MouseWheel(100);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var reaminder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget2.Composite.McTextWidget;
remainAmount = reaminder.getText().OleValue.toString();
remainAmount=remainAmount.replace("-","");
var amountIncluTax = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite4.McTextWidget;
if(remainAmount!="0.00"){ 
  amountIncluTax.setText(remainAmount);
}
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl3;
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Delay(4000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
paymentMode = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
Sys.HighlightObject(paymentMode)
paymentMode = paymentMode.getText().OleValue.toString();
Log.Message("paymentMode :"+paymentMode);
ExlAmount = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
ExlAmount = ExlAmount.getText().OleValue.toString();
Log.Message("ExlAmount :"+ExlAmount);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management",paymentMode)
ExcelUtils.WriteExcelSheet("VendorInvoice Amount",EnvParams.Opco,"Data Management",remainAmount);
ExcelUtils.WriteExcelSheet("VendorInvoice ExclAmount",EnvParams.Opco,"Data Management",ExlAmount);
TextUtils.writeLog("Tax is Validated");
TextUtils.writeLog("Tax Details is Entered and Saved");
var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
action.Click();
aqUtils.Delay(2000, "Waiting for Action");
action.Click();
//aqUtils.Delay(2000, "Waiting for Action");
//aqUtils.Delay(3000, Indicator.Text);
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
  Delay(4000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var action = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.GroupToolItemControl;
  WorkspaceUtils.waitForObj(action);
  action.Click();
  aqUtils.Delay(2000, "Waiting for Action");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(EnvParams.Country.toUpperCase()=="INDIA")
  Runner.CallMethod("IND_VendorInvoice.InvoiceSubmit",action);
  else{
    
  action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit for Approval").OleValue.toString().trim());
  }
  ReportUtils.logStep_Screenshot();
  aqUtils.Delay(8000, "Submitted for Approval");;
  TextUtils.writeLog("Invoice is Submitted for Approval");
  /*
  action.Click();
  Delay(2000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  ReportUtils.logStep_Screenshot();
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  */
  
  
//aqUtils.Delay(5000, Indicator.Text);

var journalNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget.getText().OleValue.toString().trim();
ValidationUtils.verify(true,true,"Created Vendor Invoice Journal Number :"+journalNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Invoice Journal NO",EnvParams.Opco,"Data Management",journalNumber);
ExcelUtils.WriteExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management",InvoiceNo);
ExcelUtils.WriteExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management",dueDate)
ExcelUtils.WriteExcelSheet("Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management",paymentMode)
ExcelUtils.WriteExcelSheet("VendorInvoice Amount",EnvParams.Opco,"Data Management",remainAmount);
ExcelUtils.WriteExcelSheet("VendorInvoice ExclAmount",EnvParams.Opco,"Data Management",ExlAmount);

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
    code.setText("Vendor Invoice");
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
      if((table.getItem(i).getText_2(0).OleValue.toString().trim()=="Vendor Invoice")&&(table.getItem(i).getText_2(1).OleValue.toString().trim()==EnvParams.Opco)){ 
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
