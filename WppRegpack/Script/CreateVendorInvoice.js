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
var Ipreparation_ID,IpreparationUnit,JobNoTo = "";
var company,PurchOrderNo,InvoiceNo,Description,IDate,EDate,TDSValue,PO_Value,Tax_Amount,Transaction_No ="";
 var DelimitedFile = "";
 var Model_Excel,Hitpoint_Password,Hitpoint_User_Name,URL,Directory ="";
var Language = "";
var paymentMode = "";
var ExlAmount = "";
var remainAmount = "";
var Final_Time,FinalFile,temp,File_Name = "";

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
company,PurchOrderNo,InvoiceNo,Description,IDate,EDate,TDSValue,PO_Value,Tax_Amount,Transaction_No ="";
Model_Excel,Hitpoint_Password,Hitpoint_User_Name,URL,Directory ="";
DelimitedFile = "";
paymentMode = "";
ExlAmount = "";
remainAmount = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
//try{
getDetails();

if(Hitpoint.toUpperCase()=="YES"){
goto_Purchase_Order();
getting_PO_Details();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
Hitpoint_Integration();
FileFinder();
Filemodify();
goTo_Account_Payable();
Import_Vendor_Invoice();
Submit_Vendor_Invocie();

}else{ 
goTo_Account_Payable(); 
invoiceAllocation();
}
//}catch(err){ 
//  Log.Message(err);
//}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

function goTo_Account_Payable(){ 
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

Description = ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
if((Description==null)||(Description=="")){ 
ValidationUtils.verify(false,true,"Description is Needed to Create a Vendor Invoice");
}


Hitpoint = ExcelUtils.getColumnDatas("AP Fapio Hitpoint",EnvParams.Opco)
if((EnvParams.Country.toUpperCase()=="CHINA")||(EnvParams.Country.toUpperCase()=="HONG KONG")){
if((Hitpoint==null)||(Hitpoint=="")){ 
ValidationUtils.verify(false,true,"AP Fapio Hitpoint YES/NO is Needed to Create a Vendor Invoice");
}
}

if(Hitpoint.toUpperCase()=="YES"){
  
  Transaction_No = ExcelUtils.getColumnDatas("Transaction Number",EnvParams.Opco)
  if((Transaction_No==null)||(Transaction_No=="")){ 
  ValidationUtils.verify(false,true,"Transaction_No is Needed to Create a Vendor Invoice");
  }
}

EDate = ExcelUtils.getColumnDatas("Entry Date",EnvParams.Opco)
if(Hitpoint.toUpperCase()=="YES"){
EDate = getDateFormat(0)
}
else if(EDate == "AUTOFILL")
        EDate = getSpecificDate(0);
if((EDate==null)||(EDate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Vendor Invoice");
}
//Log.Message(EDate)
IDate = ExcelUtils.getColumnDatas("Invoice Date",EnvParams.Opco)
if(Hitpoint.toUpperCase()=="YES"){
IDate = getDateFormat(0)
}
else if(IDate == "AUTOFILL")
        IDate = getSpecificDate(0);
if((IDate==null)||(IDate=="")){ 
ValidationUtils.verify(false,true,"Invoice Date is Needed to Create a Vendor Invoice");
}
//Log.Message(IDate)

if(Hitpoint.toUpperCase()=="YES"){
ExcelUtils.setExcelName(workBook, "HitPoint", true);
 Directory = ExcelUtils.getRowDatas("Download Directory","Value")
if((Directory==null)||(Directory=="")){ 
ValidationUtils.verify(false,true,"Download Directory is Needed to Create a Vendor Invoice");
}

 URL = ExcelUtils.getRowDatas("HitPoint URL","Value")
if((URL==null)||(URL=="")){ 
ValidationUtils.verify(false,true,"Hitpoint URL is Needed to Create a Vendor Invoice");
}

 Hitpoint_User_Name = ExcelUtils.getRowDatas("User Name","Value")
if((Hitpoint_User_Name==null)||(Hitpoint_User_Name=="")){ 
ValidationUtils.verify(false,true,"Hitpoint_User_Name is Needed to Create a Vendor Invoice");
}

 Hitpoint_Password = ExcelUtils.getRowDatas("Password","Value")
if((Hitpoint_Password==null)||(Hitpoint_Password=="")){ 
ValidationUtils.verify(false,true,"Hitpoint_Password is Needed to Create a Vendor Invoice");
}

 Hitpoint_Password = ExcelUtils.getRowDatas("Password","Value")
if((Hitpoint_Password==null)||(Hitpoint_Password=="")){ 
ValidationUtils.verify(false,true,"Hitpoint_Password is Needed to Create a Vendor Invoice");
}


 Model_Excel = ExcelUtils.getRowDatas("Model Delimited File","Value")
if((Model_Excel==null)||(Model_Excel=="")){ 
ValidationUtils.verify(false,true,"Model Delimited File is Needed to Create a Vendor Invoice");
}

}

TDSValue = ExcelUtils.getColumnDatas("TDS",EnvParams.Opco)
}


function invoiceAllocation(dependency){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000,"Maconomy is collection datas")
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
    
  }
 
  var dueDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.SWTObject("Composite", "", 5).SWTObject("McDatePickerWidget", "", 2);
  dueDate = dueDate.getText().OleValue.toString().trim();
  if(EnvParams.Country.toUpperCase()=="INDIA"){
  Runner.CallMethod("IND_VendorInvoice.TDS",TDSValue); 
  
  }

npEdit.MouseWheel(100);  
    
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


if(dependency){ 
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Second Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management",paymentMode)
ExcelUtils.WriteExcelSheet("Second VendorInvoice Amount",EnvParams.Opco,"Data Management",remainAmount);
ExcelUtils.WriteExcelSheet("Second VendorInvoice ExclAmount",EnvParams.Opco,"Data Management",ExlAmount);   

}else{ 
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management",paymentMode)
ExcelUtils.WriteExcelSheet("VendorInvoice Amount",EnvParams.Opco,"Data Management",remainAmount);
ExcelUtils.WriteExcelSheet("VendorInvoice ExclAmount",EnvParams.Opco,"Data Management",ExlAmount); 

}
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
if(dependency){ 
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Second Invoice Journal NO",EnvParams.Opco,"Data Management",journalNumber);
ExcelUtils.WriteExcelSheet("Second Vendor Invoice NO",EnvParams.Opco,"Data Management",InvoiceNo);
ExcelUtils.WriteExcelSheet("Second Vendor Invoice Due Date",EnvParams.Opco,"Data Management",dueDate)
ExcelUtils.WriteExcelSheet("Second Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management",paymentMode)
ExcelUtils.WriteExcelSheet("Second VendorInvoice Amount",EnvParams.Opco,"Data Management",remainAmount);
ExcelUtils.WriteExcelSheet("Second VendorInvoice ExclAmount",EnvParams.Opco,"Data Management",ExlAmount);
}else{
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Invoice Journal NO",EnvParams.Opco,"Data Management",journalNumber);
ExcelUtils.WriteExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management",InvoiceNo);
ExcelUtils.WriteExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management",dueDate)
ExcelUtils.WriteExcelSheet("Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management",paymentMode)
ExcelUtils.WriteExcelSheet("VendorInvoice Amount",EnvParams.Opco,"Data Management",remainAmount);
ExcelUtils.WriteExcelSheet("VendorInvoice ExclAmount",EnvParams.Opco,"Data Management",ExlAmount);
}
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


//Moving to Purchase Order
function goto_Purchase_Order(){ 
  
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




function getting_PO_Details(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}


//  var allPurchase = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
//  WorkspaceUtils.waitForObj(allPurchase);
//  allPurchase.Click();

var All_Purchase_Order = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(All_Purchase_Order);
All_Purchase_Order.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var companyNo = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  WorkspaceUtils.waitForObj(companyNo);
  companyNo.Click();
  companyNo.Keys("[Tab]");
  var purchaseNo = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  WorkspaceUtils.waitForObj(purchaseNo);
  purchaseNo.Click();
  purchaseNo.setText(PurchOrderNo);
var table = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
//var table = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(3000, "Reading Table Data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==PurchOrderNo){ 
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
//   var closefilter = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
   var closefilter = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
  WorkspaceUtils.waitForObj(closefilter);
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();
aqUtils.Delay(3000, "Reading Table Data");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  var PageDown = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  Sys.HighlightObject(PageDown);
  PageDown.Click();
  PageDown.MouseWheel(-15);
  aqUtils.Delay(3000, "Reading Table Data");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  PO_Value,Tax_Amount
  PO_Value = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  Sys.HighlightObject(PO_Value);
  PO_Value = PO_Value.getText().OleValue.toString().trim();
  PO_Value = PO_Value.replace(/,/g, "");
  Log.Message(PO_Value);
  
  Tax_Amount = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget;
  Sys.HighlightObject(Tax_Amount);
  Tax_Amount = Tax_Amount.getText().OleValue.toString().trim();
  Tax_Amount = Tax_Amount.replace(/,/g, "");
  Log.Message(Tax_Amount);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  }
  }
  
  
  
  
  function Hitpoint_Integration(){ 
  
//var PurchOrderNo = "1221100054";
//var InvoiceNo = "1221210100";
//var PO_Value = "7500.00";
//var Tax_Amount = "450.00";
//Browsers.Item(btChrome).Run("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201");

//ExcelUtils.setExcelName(workBook, "HitPoint", true);



Browsers.Item(btChrome).Run(URL);

var page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login");
var User_Name = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(1).Textbox("userName");
User_Name.SetText(Hitpoint_User_Name);
var Password = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(2).PasswordBox("password")
Password.SetText(Hitpoint_Password);

var CaptchaVar = BuiltIn.InputBox("Captcha", "Please enter a CAPTCHA showing in the browser", "");
var Captcha = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(3).Textbox("verifyCodeId");
Captcha.SetText(CaptchaVar);

var Login = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(4).Button("btnSysLogin");
Login.Click();

page.Wait();


 // Finds the Link object on the page
  var link = page.NativeWebObject.Find("href", "http://101.231.221.6:8090/wppoutput/base/sys/main#1", "A");

  // If the link is found
  if (link.Exists)
  {
    // Clicks the link
    link.Click();
    // Waits until the target page is loaded
    page.Wait();
  }
  
var Switch_To_Chinese = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(0).Panel(0).Panel(0).Panel(1).Link("tms_sys_select_Language_a_zh_CN");
Switch_To_Chinese.Click();

var entity = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(0).Panel(0).Panel(0).Panel(1).Link("tms_sys_select_stxx");         
entity.Click();

var Company_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel("dynamic_criteria_form_sys_selectst").Panel("querydivid").Panel(0).Form("criteriaQueryForm_sys_selectst").Table(0).Cell(0, 5).Textbox("easyui_textbox_input4");
Company_Number.Click();

Company_Number.SetText("1221");

var search = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel("dynamic_criteria_form_sys_selectst").Panel(0).Link("btn_criteriaForm_select_sys_selectst").TextNode(1);
search.Click();
aqUtils.Delay(5000, "Waiting to load");

var table = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1);
var TableSize = table.ChildCount;
Log.Message(TableSize)

var OpCo_Status = false;
for (var i=1;i<TableSize;i++){ 
Log.Message(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(i).Table(0).Cell(0, 4).Panel(0).contentText)
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(i).Table(0).Cell(0, 4).Panel(0).contentText=="1221"){ 
  var Check_Box = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(i).Table(0).Cell(0, 0).Panel(0).Checkbox("ck");
  Sys.HighlightObject(Check_Box);
  Check_Box.Click();
  aqUtils.Delay(3000, "Waiting to load");
  var OKay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel(1).Link("tms_sys_stxx_dlg_confirm").TextNode(1);
  Sys.HighlightObject(OKay);
  OKay.Click();
  OpCo_Status = true;
  var Pop_Up_Message = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel(1).Panel(1);
  Log.Message(Pop_Up_Message.contentText)
  var Pop_Up_Okay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel(2).Link(0).TextNode(0);
  Pop_Up_Okay.Click();

break;
}

}

if(!OpCo_Status){ 
  Log.Error("OpCo is not available in entity");
}

var Invoice_Warehouse = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_7").TextNode(0);
Sys.HighlightObject(Invoice_Warehouse);
//Invoice_Warehouse.Click();

var VAT_Special_Invoice = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_8").TextNode(0);
Sys.HighlightObject(VAT_Special_Invoice);
VAT_Special_Invoice.Click();
aqUtils.Delay(3000, "Waiting to load");

//                     Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(0, 0).Panel(0).Checkbox("ck")
var First_CheckBox = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(0, 0).Panel(0).Checkbox("ck");
Sys.HighlightObject(First_CheckBox);
First_CheckBox.Click();
aqUtils.Delay(8000, "Waiting to load");
var Edit_Invoice = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel("dynamic_criteria_toolbar_102001").Panel("buttondivid").Link("dynamic_criteria_button_10200102");
Sys.HighlightObject(Edit_Invoice);
Edit_Invoice.Click();
aqUtils.Delay(8000, "Waiting to load");
//Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(15).Panel(1).Panel(1)
//Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(15).Panel(2).Link(0).TextNode(0)

  var Pop_Up_Message = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(17).Panel(1).Panel(1);
  Log.Message(Pop_Up_Message.contentText)
  var Pop_Up_Okay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(17).Panel(2).Link(0);
  Pop_Up_Okay.Click();
aqUtils.Delay(8000, "Waiting to load");
  var Invoice_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(14).Panel("dynamic_criteria_dialog_102001").Panel(0).Form("fp_edit_dlg_form_104001_1").Table(0).Cell(0, 0).Fieldset(0).Panel(1).Panel(1).Textbox("easyui_textbox_input21");
  Sys.HighlightObject(Invoice_Number);
  Invoice_Number.Click();
  Invoice_Number.SetText(InvoiceNo)
  
  var Total_Amount = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(14).Panel("dynamic_criteria_dialog_102001").Panel(0).Form("fp_edit_dlg_form_104001_1").Table(0).Cell(0, 0).Fieldset(0).Panel(3).Panel(1).Textbox("easyui_textbox_input22");
  Sys.HighlightObject(Total_Amount);
  Total_Amount.Click();
  Total_Amount.SetText(PO_Value)
  
  var TaxAmount = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(14).Panel("dynamic_criteria_dialog_102001").Panel(0).Form("fp_edit_dlg_form_104001_1").Table(0).Cell(0, 0).Fieldset(0).Panel(4).Panel(1).Textbox("easyui_textbox_input23");
  Sys.HighlightObject(TaxAmount);
  TaxAmount.Click();
  TaxAmount.SetText(Tax_Amount)
  
  var Remarks = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(14).Panel("dynamic_criteria_dialog_102001").Panel(0).Form("fp_edit_dlg_form_104001_1").Table(0).Cell(0, 0).Fieldset(0).Panel(11).Panel(1).Textbox("easyui_textbox_input30");
  Sys.HighlightObject(Remarks);
  Remarks.Click();
  Remarks.SetText(PurchOrderNo)
  
  var PO = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(14).Panel("dynamic_criteria_dialog_102001").Panel(0).Form("fp_edit_dlg_form_104001_1").Table(0).Cell(0, 0).Fieldset(0).Panel(16).Panel(1).Textbox("easyui_textbox_input35");
  Sys.HighlightObject(PO);
  PO.Click();
  PO.SetText(PurchOrderNo)
  
  var Save = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(14).Panel(1).Link("sinv_fpEditDetail_dlg_save_104002");
  Sys.HighlightObject(Save);
  Save.Click();
  
  Pop_Up_Message = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001").Panel(15).Panel(1).Panel(1)
  Log.Message(Pop_Up_Message.contentText)
  Pop_Up_Okay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001").Panel(15).Panel(2).Link(0).TextNode(0)
  Pop_Up_Okay.Click();
  aqUtils.Delay(5000, "Waiting to load");
  
  var Invoice_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel("dynamic_criteria_form_102001").Panel("querydivid").Panel(0).Form("criteriaQueryForm_102001").Table(0).Cell(1, 1).Textbox("easyui_textbox_input10");
  Sys.HighlightObject(Invoice_Number);
  Invoice_Number.Click();
  Invoice_Number.SetText(InvoiceNo);
  
  var Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel("dynamic_criteria_form_102001").Panel(1).Link("btn_criteriaForm_select_102001");
  Sys.HighlightObject(Query);
  Query.Click();
  aqUtils.Delay(5000, "Waiting to load");
  var table = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0);
  var TableSize = table.ChildCount;
Log.Message(TableSize)

var Invoice_Status = false;
for (var i=0;i<TableSize;i++){ 
Log.Message(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 3).Panel(0).contentText)
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 3).Panel(0).contentText==InvoiceNo){ 
  var Check_Box = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102001").Panel("dynamic_criteria_table_area_102001").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 0).Panel(0).Checkbox("ck");
  Sys.HighlightObject(Check_Box);
  Check_Box.Click();
  aqUtils.Delay(3000, "Waiting to load");
  Invoice_Status = true;
  break;
  }
  }
  
  if(!Invoice_Status){ 
    Log.Error("Invoice is Not Created");
    }
  var Comprehensive_Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102001", 0).Panel(1).Panel(0).Panel(0).Panel("easyui_tree_11").TextNode(0);
  Sys.HighlightObject(Comprehensive_Query);
  Comprehensive_Query.Click();
  
  var Invoice_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel("dynamic_criteria_form_102005").Panel("querydivid").Panel(0).Form("criteriaQueryForm_102005").Table(0).Cell(1, 3).Textbox("easyui_textbox_input44");
    Sys.HighlightObject(Invoice_Number);
  Invoice_Number.Click();
  Invoice_Number.SetText(InvoiceNo);
  
  var Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel("dynamic_criteria_form_102005").Panel(1).Link("btn_criteriaForm_select_102005");
  Sys.HighlightObject(Query);
  Query.Click();
  aqUtils.Delay(5000, "Waiting to load");
  Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(0, 7).Panel(0)
  var table = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0);
  var TableSize = table.ChildCount;
Log.Message(TableSize)

var Invoice_Status = false;
for (var i=0;i<TableSize;i++){ 
Log.Message(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 7).Panel(0).contentText)
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 7).Panel(0).contentText==InvoiceNo){ 
  var Check_Box = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 0).Panel(0).Checkbox("ck");
  Sys.HighlightObject(Check_Box);
  Check_Box.Click();
  aqUtils.Delay(3000, "Waiting to load");
  
  var Export_Invoice_Data = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&102005", 0).Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_102005").Panel("dynamic_criteria_table_area_102005").Panel("dynamic_criteria_toolbar_102005").Panel("buttondivid").Link("dynamic_criteria_button_10200508");
  Sys.HighlightObject(Export_Invoice_Data);
  Export_Invoice_Data.Click();
  
  Invoice_Status = true;
  break;
  }
  }
  
  if(!Invoice_Status){ 
    Log.Error("Invoice is Not Created");
    }

  
  
}



function FileFinder()
{
  
aqUtils.Delay(60000,"Checking the downloads files")
  var foundFiles, aFile;
  var First = true
  var sPath = "";
  Final_Time,FinalFile,temp,File_Name = "";
  


// Creates a file
aqFile.Create(sPath);
foundFiles = aqFileSystem.FindFiles(Directory, "*.txt");
if (!strictEqual(foundFiles, null))
    while (foundFiles.HasNext())
    {
      aFile = foundFiles.Next();
//      Log.Message(aFile.Name);
      
      if((First)&&(aFile.Name.indexOf("发票数据信息")!=-1)){ 
        var sPath = Directory+aFile.Name;
        var FileInf = aqFileSystem.GetFileInfo(sPath);
         Final_Time = FileInf.DateLastModified;
         FinalFile = sPath;
         File_Name = aFile.Name;
         First = false;
      }
      else{ 
      if(aFile.Name.indexOf("发票数据信息")!=-1){ 
        var sPath = Directory+aFile.Name
        var FileInf = aqFileSystem.GetFileInfo(sPath);
        var Time = FileInf.DateLastModified;
        temp = sPath;
       
       if(Time>Final_Time){ 
       Final_Time = FileInf.DateLastModified;
       FinalFile = sPath;
       File_Name = aFile.Name;
       }
      }
      
      }
    }
  else
    Log.Message("No files were found.");
    
    Log.Message("Final File Path")
    Log.Message(FinalFile);
    Log.Message("File Modified Time");
    Log.Message(Final_Time)
    Log.Message("File Name");
    Log.Message(File_Name)
    
}


function Filemodify(){ 
Log.Message(FinalFile)
//var Model_Excel = "C:\\Users\\674087\\Music\\Final Global Scripts Making\\WppRegpack\\TestResource\\AP FAPIO Import Files.xlsx"
// var Return = JavaClasses.Text_Modify.Edit_AP_Fapio.change_Text(FinalFile,"#KEEP	#KEEP	"+Transaction_No,EDate,FinalFile);
// Log.Message(Return);
 DelimitedFile = JavaClasses.Text_Modify.Text_Delimited.change_Text(FinalFile,Model_Excel,Transaction_No,IDate,EDate,"")
 
  DelimitedFile = Model_Excel;
  DelimitedFile = DelimitedFile.substring(0,DelimitedFile.indexOf(".xlsx"));
  DelimitedFile = DelimitedFile +"_"+EnvParams.Opco+".txt";
 Log.Message("DelimitedFile : "+DelimitedFile);
 
}

  function Import_Vendor_Invoice(){ 
    aqUtils.Delay(4000,"Import Files")
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    aqUtils.Delay(2000,"Import Files")
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    
    var Invoice_Allocation = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(Invoice_Allocation);
    Invoice_Allocation.Click();
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    var CloseFilter = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(CloseFilter);
    CloseFilter.Click();
    aqUtils.Delay(4000,"Clicking Actions")
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    aqUtils.Delay(4000,"Clicking Actions")
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    
    var Action = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.GroupToolItemControl;
    Sys.HighlightObject(Action);
    Action.Click();
    aqUtils.Delay(5000,"Clicking Actions")
    Action.Click();
    //aqUtils.Delay(3000, Indicator.Text);
    Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Invoices").OleValue.toString().trim());
    ReportUtils.logStep_Screenshot();
    aqUtils.Delay(8000,"Clicking Actions")
    var Internal_Names = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Internal_Names);
    if(!Internal_Names.getSelection()){
      Internal_Names.Click();
      ValidationUtils.verify(true,true,"Internal_Names is Checked");
    } 
    
    var Progress_Bar = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Progress_Bar);
    if(!Progress_Bar.getSelection()){
      Progress_Bar.Click();
      ValidationUtils.verify(true,true,"Progress_Bar is Checked");
    } 
    var Logging = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Logging);
    if(!Logging.getSelection()){
      Logging.Click();
      ValidationUtils.verify(true,true,"Logging is Checked");
    }
    var Echo = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Echo);
    if(Echo.getSelection()){
      Echo.Click();
      ValidationUtils.verify(true,true,"Echo is UnChecked");
    }
    var Help = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Help);
    if(Help.getSelection()){
      Help.Click();
      ValidationUtils.verify(true,true,"Help is UnChecked");
    }
    var Report_Error_Lines = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Report_Error_Lines);
    if(Report_Error_Lines.getSelection()){
      Report_Error_Lines.Click();
      ValidationUtils.verify(true,true,"Report_Error_Lines is UnChecked");
    }
    var Print_Log = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    Sys.HighlightObject(Print_Log);
    if(Print_Log.getSelection()){
      Print_Log.Click();
      ValidationUtils.verify(true,true,"Print_Log is UnChecked");
    }
    
    var Import_Invoice = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Vendor Invoices").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Invoices").OleValue.toString().trim());
    Sys.HighlightObject(Import_Invoice);
    Import_Invoice.Click();
    
    Log.Message("FinalFile :"+DelimitedFile)
    aqUtils.Delay(4000, "Waiting to Open file");;
    var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    dicratory.SetText(FinalFile);
    dicratory.SetText(DelimitedFile);
    var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
    Sys.HighlightObject(opendoc);
    opendoc.HoverMouse();
    ReportUtils.logStep_Screenshot();
    opendoc.Click();
    aqUtils.Delay(2000, "Document Attached");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
    aqUtils.Delay(2000, "Document Attached");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }

    
    var Save_Import_Message = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save file").OleValue.toString().trim(), 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    Save_Import_Message.Keys(workBook);
    aqUtils.Delay(2000, Indicator.Text);
    var SaveTitle = Save_Import_Message.wText;
    
    sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
    if (! aqFileSystem.Exists(sFolder)){
    if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
    }
    else{
    Log.Error("Could not create the folder " + sFolder);
    }
    }
    Save_Import_Message.Keys(sFolder+SaveTitle);

    var Save = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save file").OleValue.toString().trim(), 1).Window("Button", "&Save", 1)
    Save.Click();
    aqUtils.Delay(5000, "Document Attached");

var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save file").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
    var SaveAs = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Save file").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
    SaveAs.Click();
}
    aqUtils.Delay(2000, "Document Attached");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
    
    
     var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Import Vendor Invoices").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Import Vendor Invoices").OleValue.toString().trim(), 2000);
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
  

function Submit_Vendor_Invocie(){ 
  
aqUtils.Delay(4000,"Selecting Vendor Invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var ShowFilter = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(ShowFilter);
ShowFilter.Click();
aqUtils.Delay(4000,"Selecting Vendor Invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var company = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(company);
company.Click();
company.setText(EnvParams.Opco);
aqUtils.Delay(2000,"Selecting Vendor Invoice");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Selecting Vendor Invoice");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Selecting Vendor Invoice");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Selecting Vendor Invoice");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Selecting Vendor Invoice");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Selecting Vendor Invoice");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);

var Invoice = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Invoice);
Invoice.Click();
Invoice.setText(InvoiceNo);
aqUtils.Delay(5000,"Selecting Vendor Invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000,"Selecting Vendor Invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;

  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(6).OleValue.toString().trim()==InvoiceNo){ 
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
  ValidationUtils.verify(flag,true,"Created Vendor Invoice is available in system");
  TextUtils.writeLog("Created Vendor Invoicer is available in system");
  
  if(flag){ 
var closeFilter = Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(closeFilter);
closeFilter.Click();
aqUtils.Delay(3000,"Selecting Vendor Invoice");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
 
  var dueDate = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.SWTObject("Composite", "", 5).SWTObject("McDatePickerWidget", "", 2);
  dueDate = dueDate.getText().OleValue.toString().trim()
  
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

    
  action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit for Approval").OleValue.toString().trim());
  ReportUtils.logStep_Screenshot();
  aqUtils.Delay(8000, "Submitted for Approval");;
  TextUtils.writeLog("Invoice is Submitted for Approval");


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
  

}



// Creating 2nd Vendor Invoice for Reverse invoice, Credit Note, Reverse Credit Note
function CreateInvoice_Dependency(){ 
TextUtils.writeLog("Create Vendor Invoice Started"); 
Indicator.PushText("waiting for window to open");

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "Dependency_VendorInvoice";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
company,PurchOrderNo,InvoiceNo,Description,IDate,EDate,TDSValue,PO_Value,Tax_Amount,Transaction_No ="";
Model_Excel,Hitpoint_Password,Hitpoint_User_Name,URL,Directory ="";
DelimitedFile = "";
paymentMode = "";
ExlAmount = "";
remainAmount = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
//try{
getDetails_for_Dependency();
Create_New_Job_2();

aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

goTo_Account_Payable(); 
invoiceAllocation(true);
//}catch(err){ 
//  Log.Message(err);
//}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails_for_Dependency(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);

company = ExcelUtils.getColumnDatas("Opco",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Vendor Invoice");
}
//Log.Message(company)

ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(PurchOrderNo)
InvoiceNo = ExcelUtils.getColumnDatas("Invoice No",EnvParams.Opco)
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Invoice No. is Needed to Create a Vendor Invoice");
}
//Log.Message(InvoiceNo)

Description = ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
if((Description==null)||(Description=="")){ 
ValidationUtils.verify(false,true,"Description is Needed to Create a Vendor Invoice");
}

EDate = ExcelUtils.getColumnDatas("Entry Date",EnvParams.Opco)
if(EDate == "AUTOFILL")
        EDate = getSpecificDate(0);
if((EDate==null)||(EDate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Vendor Invoice");
}
//Log.Message(EDate)
IDate = ExcelUtils.getColumnDatas("Invoice Date",EnvParams.Opco)
if(IDate == "AUTOFILL")
        IDate = getSpecificDate(0);
if((IDate==null)||(IDate=="")){ 
ValidationUtils.verify(false,true,"Invoice Date is Needed to Create a Vendor Invoice");
}
//Log.Message(IDate)

TDSValue = ExcelUtils.getColumnDatas("TDS",EnvParams.Opco)
}



function Create_New_Job_2(){ 
  
    //Creation of Job
    Ipreparation_ID,IpreparationUnit = "";
    Ipreparation_ID = TestRunner.testCaseId;
    IpreparationUnit = TestRunner.unitName; 
    TestRunner.TempUnit = IpreparationUnit;
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var jobSheet = ExcelUtils.getColumnDatas("Job Sheet",EnvParams.Opco)
    if(jobSheet==""){ 
      ValidationUtils.verify(true,false,"Need Job to Create Vendor Invoice")
    }
    
    ExcelUtils.setExcelName(workBook, jobSheet, true);
    var serialOder = ExcelUtils.getRowDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Vendor Invoice")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    JobNoTo = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)
    if((JobNoTo=="")||(JobNoTo==null)){
      
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    Job_JIRAID = ExcelUtils.getRowDatas("JobCreation_"+serialOder,EnvParams.Country);
    if((Job_JIRAID=="")||(Job_JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for Jobcreation_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = Job_JIRAID;
    TestRunner.unitName = "JobCreation_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+Job_JIRAID)
    Runner.CallMethod("Creation_Of_Job.createJob",jobSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    JobNoTo = ExcelUtils.getRowDatas("Job Number_"+serialOder,EnvParams.Opco)  
     
    //Creation of Budget
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var budgetSheet = ExcelUtils.getColumnDatas("Budget sheet",EnvParams.Opco)
    if(budgetSheet==""){ 
      ValidationUtils.verify(true,false,"Need Working Estimate for Job to Create Vendor Invoice")
    }
    ExcelUtils.setExcelName(workBook, budgetSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Budget")
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var WE_Number = ExcelUtils.getRowDatas("Working Estimate_"+serialOder,EnvParams.Opco)
    if((WE_Number=="")||(WE_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateBudget_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateBudget_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;  
    TestRunner.unitName = "CreateBudget_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Job Budget");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Job Budget");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("BudgetCreation.createBudget",budgetSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
  //Creation of Quote 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var quoteSheet = ExcelUtils.getColumnDatas("Quote Sheet",EnvParams.Opco)
    if(quoteSheet==""){ 
      ValidationUtils.verify(true,false,"Need Client Approved Estimate for Job to Create Vendor Invoice")
    }
    ExcelUtils.setExcelName(workBook, quoteSheet, true);
    var serialOder = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create Quote")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var CE_Number = ExcelUtils.getRowDatas("Client Approved Estimate_"+serialOder,EnvParams.Opco)
    if((CE_Number=="")||(CE_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreateQuote_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreateQuote_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;    
    TestRunner.unitName = "CreateQuote_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Quote");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Quote");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("Creation_of_Quote.CreateQuote",quoteSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }

    //Creation of PO
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var POSheet = ExcelUtils.getColumnDatas("PO Sheet",EnvParams.Opco)
    if(POSheet==""){ 
      ValidationUtils.verify(true,false,"Need PO for Job to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, POSheet, true);
    var JobSO = ExcelUtils.getColumnDatas("Job Serial Order",EnvParams.Opco)
    if(JobSO==""){ 
      ValidationUtils.verify(true,false,"Need Job Serial Order to Create PO")
    }
    var PO_SO = ExcelUtils.getColumnDatas("PO Serial Order",EnvParams.Opco)
    if(PO_SO==""){ 
      ValidationUtils.verify(true,false,"Need PO Serial Order to Create PO")
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var PO_Number = ExcelUtils.getRowDatas("PO Number_"+PO_SO,EnvParams.Opco)
    if((PO_Number=="")||(PO_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("CreatePurchaseOrder_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for CreatePurchaseOrder_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;   
    TestRunner.unitName = "CreatePurchaseOrder_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Creation of Purchase Order");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Creation of Purchase Order");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("CreatePO.CreatePurchaseOrder",POSheet,JobSO,PO_SO);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    PoNumber2 = ExcelUtils.getRowDatas("PO Number_"+PO_SO,EnvParams.Opco)
    PurchOrderNo = PoNumber2;
    
  //Approving PO
    ExcelUtils.setExcelName(workBook, sheetName, true);
    var APSheet = ExcelUtils.getColumnDatas("Approve PO Sheet",EnvParams.Opco)
    if(APSheet==""){ 
      ValidationUtils.verify(true,false,"Need Approve PO Sheet for Job to Create Invoice preparation")
    }
    ExcelUtils.setExcelName(workBook, APSheet, true);
    var serialOder = ExcelUtils.getRowDatas("PO Serial Order",EnvParams.Opco)
    if(serialOder==""){ 
      ValidationUtils.verify(true,false,"Need PO Serial Order to Approve PO")
    }
    
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    var AP_Number = ExcelUtils.getRowDatas("Approved PO_"+serialOder,EnvParams.Opco)
    if((AP_Number=="")||(AP_Number==null)){
    TestRunner.JiraStat = true;
    TestRunner.JiraUpdate = true;
    var xlDriver= Project.Path+TextUtils.GetProjectValue("EnvDetailsPath");
    ExcelUtils.setExcelName(xlDriver, "JIRA_Details", true);
    var JIRAID = ExcelUtils.getRowDatas("ApprovePurchaseOrder_"+serialOder,EnvParams.Country);
    if((JIRAID=="")||(JIRAID==null)){
      ValidationUtils.verify(true,false,"JIRA ID for ApprovePurchaseOrder_"+serialOder+" is needed");
      }
    TestRunner.testCaseId = JIRAID;    
    TestRunner.unitName = "ApprovePurchaseOrder_"+serialOder;
    ReportUtils.DStat = true;
    var reportName = "Report_"+EnvParams.Opco+"_"+TestRunner.unitName;
    ReportUtils.createDependencyReport(reportName);
    ReportUtils.DependycreateTest(TestRunner.unitName, "Approve Purchase Order");
    var FolderID = Log.CreateFolder(EnvParams.Opco+"_Approve Purchase Order");
    Log.PushLogFolder(FolderID);
    Log.Message("TestCase ID: "+JIRAID)
    Runner.CallMethod("ApprovePO.ApprovePurchaseOrder",APSheet,serialOder);
    Log.PopLogFolder();
    
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
   }

TestRunner.testCaseId = Ipreparation_ID;
TestRunner.unitName = IpreparationUnit;

Log.Message("Job To :"+JobNoTo)



TestRunner.JiraStat = true;
TestRunner.JiraUpdate = true;

}
