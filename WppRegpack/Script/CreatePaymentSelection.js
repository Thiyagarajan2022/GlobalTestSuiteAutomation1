//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreatePaymentSelection";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var Duedate="";
var VendorNo="";
var Paymentagent="";
var Paymodemode="";
var ExchangeDate="";
var layoutTypes="";
var Invoicenumber="";
var amount ="";

//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
//Paymentagent = ExcelUtils.getRowDatas("Payment_Agent",EnvParams.Opco)
//Log.Message(Paymentagent)
//if((Paymentagent==null)||(Paymentagent=="")){ 
//ValidationUtils.verify(false,true,"Payment Agent is Needed to Create a Payment Selection");
//}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Paymodemode = ReadExcelSheet("Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management");
if((Paymodemode=="")||(Paymodemode==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Paymodemode = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco) 
}
Log.Message(Paymodemode)
if((Paymodemode==null)||(Paymodemode=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Duedate = ReadExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management");
if((Duedate=="")||(Duedate==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Duedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
}
Log.Message(Duedate)
if((Duedate==null)||(Duedate=="")){ 
ValidationUtils.verify(false,true,"Due Date Number is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
amount = ReadExcelSheet("VendorInvoice Amount",EnvParams.Opco,"Data Management");
if((amount=="")||(amount==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
amount = ExcelUtils.getRowDatas("Amount",EnvParams.Opco)
}
Log.Message(amount)
if((amount==null)||(amount=="")){ 
ValidationUtils.verify(false,true,"Amount is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
layoutTypes = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(layoutTypes)
if((layoutTypes==null)||(layoutTypes=="")){ 
ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Invoicenumber = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((Invoicenumber=="")||(Invoicenumber==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Invoicenumber = ExcelUtils.getRowDatas("Vendor Invoice NO",EnvParams.Opco)
}
Log.Message(Invoicenumber)
if((Invoicenumber==null)||(Invoicenumber=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice Nunber is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
Log.Message(VendorNo)
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
Log.Message(VendorNo)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
}
}




function CreatePaymentSeletion() {
ReportUtils.logStep("INFO", "Enter Bank Details");
while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

 var banking = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(banking);
  WorkspaceUtils.waitForObj(banking);
  var create = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  create.HoverMouse();
  ReportUtils.logStep_Screenshot("");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  var vendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  Sys.HighlightObject(vendor);
  if(VendorNo!=""){
  vendor.Click();
  WorkspaceUtils.VPWSearchByValue(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
//  WorkspaceUtils.VPWSearchByValue(vendor,"Vendor",VendorNo,"Vendor Number");
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
  }
  
  var vendor1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget2;
   Sys.HighlightObject(vendor1);
  if(VendorNo!=""){
  vendor1.Click();
  WorkspaceUtils.VPWSearchByValue(vendor1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
    }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
  }
  
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var company = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
    waitForObj(company);
  Sys.HighlightObject(company)
  company.Click();
  WorkspaceUtils.SearchByValue(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

  aqUtils.Delay(1000, Indicator.Text);
  var company1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget2;
  Sys.HighlightObject(company1)
  company1.Click();
  WorkspaceUtils.SearchByValue(company1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

  
//  var paymentAgent = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
//   if(Paymentagent!=""){
//  paymentAgent.Click();
//  WorkspaceUtils.SearchByValue(paymentAgent,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Agent").OleValue.toString().trim(),Paymentagent,"Payment Agent")
//}else{ 
//  ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment Selection");
//}
  
  Log.Message(Paymodemode)
  var paymentMode = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget;
   if(Paymodemode!=""){
  paymentMode.Click();
  WorkspaceUtils.SearchByValue(paymentMode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Mode").OleValue.toString().trim(),Paymodemode,"Payment Mode")
  }else{ 
    ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment Selection");
  }

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var createselection = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McPlainCheckboxView.Button;
    createselection.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  Log.Message(Duedate)
   var duedate = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McDatePickerWidget;
   Sys.HighlightObject(duedate);
   
   if(duedate.getText()!=Duedate){
      if(Duedate!=""){
       aqUtils.Delay(1000, Indicator.Text);
       duedate.setText(Duedate);
//          WorkspaceUtils.CalenderDateSelection(duedate,DueDate)
          ValidationUtils.verify(true,true,"Due Date is selected in Maconomy"); 
        }
    }
    else{ 
      ValidationUtils.verify(false,true,"Due Date is Needed to Create a Payment Selection");
    } 
  ReportUtils.logStep_Screenshot();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var scroll = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  scroll.MouseWheel(-200);
  aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


  var layout = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget4.Composite2.McPopupPickerWidget;
  Log.Message(layoutTypes)
  layout.Keys(layoutTypes);
  Delay(5000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  waitForObj(save)
  Sys.HighlightObject(save)
  save.Click();
  ReportUtils.logStep_Screenshot();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var print = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  waitForObj(print)
  print.Click();

  

while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  aqUtils.Delay(1000,"PDF file is getting generated");
}


aqUtils.Delay(6000, Indicator.Text);
WorkspaceUtils.savePDF_And_WriteToExcel("PaymentSelectionMpl","PaymentSelection");
  
}




//Go To Banking from Menu
function goToBankingMenuItem(){

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.Banking.Exists()){
 ImageRepository.ImageSet.Banking.Click();
}
else if(ImageRepository.ImageSet.Banking1.Exists()){
ImageRepository.ImageSet.Banking1.Click();
}
else{
ImageRepository.ImageSet.Banking2.Click();
}

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//Delay(3000);
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
//Client_Managt.ClickItem("|Bank Transactions");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions").OleValue.toString().trim());
//Client_Managt.DblClickItem("|Bank Transactions");
}
}
ReportUtils.logStep("INFO", "Moved to Banking Transactions from job Menu");
TextUtils.writeLog("Entering into Banking Transactions from Jobs Menu");
}


//Main Function
function CreatePayment() {
TextUtils.writeLog("Create Payment Selection Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreatePaymentSelection";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
VendorNo,Paymentagent,Paymodemode ="";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

try{
getDetails();
goToBankingMenuItem();   
CreatePaymentSeletion(); 
}
  catch(err){
    Log.Message(err);
  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


