//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT CreateVendorInvoice


/** 
 * This script reverse created vendor invoice
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :03/23/2021
 */
 
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Reverse Invoice";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var company,PurchOrderNo,InvoiceNo,Description,NewInvoiceNo,IDate,EDate,TDSValue ="";
var Language = "";
var POnum = "";
var VInum = "";

function ReverseInvoice(sheet,PO,VIno){ 
TextUtils.writeLog("Reverse Vendor Invoice Started"); 
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
sheetName = sheet;
POnum =PO;
VInum =VIno;

level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
company,PurchOrderNo,InvoiceNo,NewInvoiceNo,Description,IDate,EDate,TDSValue ="";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
try{
getDetails();
goToJobMenuItem(); 
invoiceAllocation();
}catch(err){ 
  Log.Message(err);
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

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

ExcelUtils.setExcelName(workBook, "Data Management", true);
//PurchOrderNo = ReadExcelSheet("PO Number",EnvParams.Opco,"Data Management");
//Log.Message(PurchOrderNo)
//if((PurchOrderNo=="")||(PurchOrderNo==null)){
//ExcelUtils.setExcelName(workBook, sheetName, true);
//PurchOrderNo = ExcelUtils.getColumnDatas("Purch Order No",EnvParams.Opco)
//}
//if((PurchOrderNo==null)||(PurchOrderNo=="")){ 
//ValidationUtils.verify(false,true,"PO Number is Needed to Create a Vendor Invoice");
//}

InvoiceNo = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management");
Log.Message(InvoiceNo)
if((InvoiceNo=="")||(InvoiceNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
InvoiceNo = ExcelUtils.getColumnDatas("Invoice No",EnvParams.Opco)
}
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Invoice Number is Needed to Create a Vendor Invoice");
}
sheetName = "Reverse Invoice"
ExcelUtils.setExcelName(workBook, sheetName, true);
EDate = ExcelUtils.getColumnDatas("Entry Date",EnvParams.Opco)
Log.Message(EDate)
if((EDate==null)||(EDate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Vendor Invoice");
}
//Log.Message(EDate)
IDate = ExcelUtils.getColumnDatas("Invoice Date",EnvParams.Opco)
Log.Message(IDate)
if((IDate==null)||(IDate=="")){ 
ValidationUtils.verify(false,true,"Invoice Date is Needed to Create a Vendor Invoice");
}
NewInvoiceNo = ExcelUtils.getColumnDatas("New Invoice No",EnvParams.Opco)
Log.Message(NewInvoiceNo)
if((NewInvoiceNo==null)||(NewInvoiceNo=="")){ 
ValidationUtils.verify(false,true,"New Invoice No is Needed to Create a Vendor Invoice");
}
//Log.Message(IDate)
Description = ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
Log.Message(Description)
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
TextUtils.writeLog("New Invoice Button is Clicked");


aqUtils.Delay(2000, "Waiting for Action");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(5000, "Waiting for Action");
var Create_Method = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McPopupPickerWidget", "", 2);
Create_Method.Keys(" ");
aqUtils.Delay(5000, "Waiting for Action");
Create_Method.Click();
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Invoice").OleValue.toString().trim(),"Create Method");
aqUtils.Delay(2000, "Waiting for Action");

var Next = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
Sys.HighlightObject(Next);
Next.Click();
aqUtils.Delay(5000, "Waiting for Action");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//var Vendorno = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);


var invoicenumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
Sys.HighlightObject(invoicenumber)
 if(InvoiceNo!=""){
  invoicenumber.Click();
//  VPWSearchByValue(invoicenumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Invoice").OleValue.toString().trim(),InvoiceNo,"Invoice Number");
  VPWSearchByValue(invoicenumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Invoice").OleValue.toString().trim(),InvoiceNo,"Invoice Number");
  TextUtils.writeLog("Vendor is selected from macanomy:"+InvoiceNo+"");  
  }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment File");
  }
  
var ReverseCopying = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
if(ReverseCopying.getSelection()){ 
  ReverseCopying.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  ReportUtils.logStep("INFO", "ReverseCopying");
    Log.Message("ReverseCopying")
    checkmark = true;
  }
  else{
    ReverseCopying.Click();
    TextUtils.writeLog("ReverseCopying is Clicked");
  }
  
  
var OriginalExchangeRate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//if(OriginalExchangeRate.getSelection()){ 
//  OriginalExchangeRate.HoverMouse();
//  ReportUtils.logStep_Screenshot("");
//  ReportUtils.logStep("INFO", "OriginalExchangeRate");
//    Log.Message("OriginalExchangeRate")
//    checkmark = true;
//  }
//  else{
//    OriginalExchangeRate.Click();
//    TextUtils.writeLog("OriginalExchangeRate is Clicked");
//  }
  
var EntryDate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2);
if(EDate!=""){
  EntryDate.Click();
EntryDate.setText(EDate);
  }
  
var invoiceDate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2);
if(IDate!=""){
invoiceDate.setText(IDate);
  }
  
var Descrip = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
if(Description!=""){
  Descrip.Click();
Descrip.setText(Description);
ValidationUtils.verify(true,true,"Description Entered in Invoice Allocation");
   }
   
aqUtils.Delay(5000, "Waiting for Action");
var InvoiceType = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2);
InvoiceType.Keys(" ");
aqUtils.Delay(5000, "Waiting for Action");
InvoiceType.Click();
Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim())
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim(),"Create Method");
aqUtils.Delay(2000, "Waiting for Action");

//InvoiceType.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Memo").OleValue.toString().trim());
aqUtils.Delay(5000, "Waiting for Action");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(EnvParams.Country.toUpperCase()=="INDIA"){  

var TransactionType =  Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McValuePickerWidget", "", 2)
//if((TransactionType.getText()=="")||(TransactionType.getText()==null)){
TransactionType.Click();
SearchByValue(TransactionType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transaction Type").OleValue.toString().trim(),"Transaction Type");
//}  
}

var InvoiceNumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
if(InvoiceNo!=""){
InvoiceNumber.setText(NewInvoiceNo);
ValidationUtils.verify(true,true,"Invoice No Entered in Invoice Allocation");
   }
   
   
var companyNo = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(EnvParams.Opco!=""){
companyNo.Click();
WorkspaceUtils.SearchByValue(companyNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
  }
var Create = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Vendor Invoice").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
Sys.HighlightObject(Create);
Create.Click();

  
//TextUtils.writeLog("New Invoice Button is Clicked");
/*
var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
if(EnvParams.Opco!=""){
companyNo.Click();
WorkspaceUtils.SearchByValue(companyNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
  }

  Delay(5000);
if(EnvParams.Country.toUpperCase()=="INDIA"){    
var TransactionType = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget3.Composite.McValuePickerWidget;
if((TransactionType.getText()=="")||(TransactionType.getText()==null)){
TransactionType.Click();
SearchByValue(TransactionType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transaction Type").OleValue.toString().trim(),"Transaction Type");
  }
}

var InvoiceType = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite6.McPopupPickerWidget;
InvoiceType.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice").OleValue.toString().trim());

var EntryDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite3.McDatePickerWidget;
if(EDate!=""){
  EntryDate.Click();
EntryDate.setText(EDate);
  }
var invoiceDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite5.McDatePickerWidget;
if(IDate!=""){
invoiceDate.setText(IDate);
  }
  
var InvoiceNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.McTextWidget;
if(InvoiceNo!=""){
InvoiceNumber.setText(NewInvoiceNo);
ValidationUtils.verify(true,true,"Invoice No Entered in Invoice Allocation");
   }
     
var Descrip = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite2.McTextWidget;
if(Description!=""){
  Descrip.Click();
Descrip.setText(Description);
ValidationUtils.verify(true,true,"Description Entered in Invoice Allocation");
   }
   
   
   
var invoicenumber = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite3.copyinvoice;
Sys.HighlightObject(invoicenumber)
 if(InvoiceNo!=""){
  invoicenumber.Click();
//  VPWSearchByValue(invoicenumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Invoice").OleValue.toString().trim(),InvoiceNo,"Invoice Number");
  VPWSearchByValue(invoicenumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Invoice").OleValue.toString().trim(),InvoiceNo,"Invoice Number");
  TextUtils.writeLog("Vendor is selected from macanomy:"+InvoiceNo+"");  
  }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment File");
  }



//invoicenumber.Click();
//invoicenumber.setText(InvoiceNo);
aqUtils.Delay(5000, "Waiting for Invoice Allocation");

var ReverseCopying = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.McPlainCheckboxView.ReverseCopyButton;
if(ReverseCopying.getSelection()){ 
  ReverseCopying.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  ReportUtils.logStep("INFO", "ReverseCopying");
    Log.Message("ReverseCopying")
    checkmark = true;
  }
  else{
    ReverseCopying.Click();
    TextUtils.writeLog("ReverseCopying is Clicked");
  }


var OriginalExchangeRate = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite2.McPlainCheckboxView.OriginalERateButton;
if(OriginalExchangeRate.getSelection()){ 
  OriginalExchangeRate.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  ReportUtils.logStep("INFO", "OriginalExchangeRate");
    Log.Message("OriginalExchangeRate")
    checkmark = true;
  }
  else{
    OriginalExchangeRate.Click();
    TextUtils.writeLog("OriginalExchangeRate is Clicked");
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

/*
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
*/

}
  }
   
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
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
 taxcode2.setText(" ") ;
}

var taxcode3 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.Composite3.McValuePickerWidget;
if(tax3!=""){
if(tax3!=taxcode3.getText()){
taxcode3.Click();
WorkspaceUtils.SearchByValue(taxcode3,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "G/L Tax Code").OleValue.toString().trim(),tax3,"Tax Code 3");
}
}
else{ 
 taxcode3.setText(" ") ;
}

var Save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl3;
WorkspaceUtils.waitForObj(Save);
Save.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
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
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit for Approval").OleValue.toString().trim());
  }
  ReportUtils.logStep_Screenshot();
  aqUtils.Delay(8000, "Submitted for Approval");;
  TextUtils.writeLog("Invoice is Submitted for Approval");

  
//aqUtils.Delay(5000, Indicator.Text);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var journalNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget.getText().OleValue.toString().trim();
ValidationUtils.verify(true,true,"Created Vendor Invoice Journal Number :"+journalNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Reverse Invoice Journal NO",EnvParams.Opco,"Data Management",journalNumber);
ExcelUtils.WriteExcelSheet("Reverse Vendor Invoice NO",EnvParams.Opco,"Data Management",NewInvoiceNo);
//ExcelUtils.WriteExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management",dueDate)
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


function VPWSearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
//    aqUtils.Delay(3000, Indicator.Text);;

    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
  waitForObj(code);
  code.Click();
//    code.setText(value);
    code.Keys("[Tab][Tab][Tab]")
    var invoice = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
    invoice.Click();
    invoice.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();    
//Sys.HighlightObject(serch);
//    if(serch.isEnabled())
//  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
//    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2)
    var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
    waitForObj(OK);
    Sys.HighlightObject(table); 
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(3).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}
