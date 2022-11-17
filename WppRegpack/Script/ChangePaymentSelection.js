//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ChangePaymentSelection";
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
var layout="";
var InvoiceNo=""
var OldDuedate = ""

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




ExcelUtils.setExcelName(workBook, sheetName, true);
Duedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
Log.Message(Duedate)
if((Duedate==null)||(Duedate=="")){ 
ValidationUtils.verify(false,true,"Due Date Number is Needed to Create a Payment Selection");
}


ExcelUtils.setExcelName(workBook, "Data Management", true);
OldDuedate = ReadExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management");
if((OldDuedate=="")||(OldDuedate==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
OldDuedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
}
Log.Message(OldDuedate)
if((OldDuedate==null)||(OldDuedate=="")){ 
ValidationUtils.verify(false,true,"Due Date Number is Needed to Change a Payment Selection");
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
layout = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(layout)
if((layout==null)||(layout=="")){ 
ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
InvoiceNo = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((InvoiceNo=="")||(InvoiceNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
InvoiceNo = ExcelUtils.getRowDatas("Vendor Invoice NO",EnvParams.Opco)
}
Log.Message(InvoiceNo)
if((InvoiceNo==null)||(InvoiceNo=="")){ 
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


function goToAp(){  
  var ap = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(ap);
  WorkspaceUtils.waitForObj(ap);
  var showvendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(showvendor);
  showvendor.Click();  
  WorkspaceUtils.waitForObj(showvendor);
  var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  var company = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  company.Click();
  company.setText(EnvParams.Opco);
  company.Keys("[Tab][Tab][Tab]");
  var vendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  Sys.HighlightObject(vendor)
  vendor.setText(VendorNo);
  ReportUtils.logStep_Screenshot();
        Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
         Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);  
   
  var Duedate1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  Sys.HighlightObject(Duedate1)  
  WorkspaceUtils.waitForObj(Duedate1);
  var Duedatee = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget2.getText();
  Sys.HighlightObject(Duedatee)
  DueDate = Duedatee.getText();
  ReportUtils.logStep_Screenshot();
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("DueDate",EnvParams.Opco,"Data Management",Duedatee)
  
}



function CreatePaymentSeletion() {
ReportUtils.logStep("INFO", "Enter Bank Details");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var banking = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(banking);
//  WorkspaceUtils.waitForObj(banking);
  var create = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
 WorkspaceUtils.waitForObj(create);
 create.Click();

  ReportUtils.logStep_Screenshot("");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var vendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  Sys.HighlightObject(vendor);
  if(VendorNo!=""){
  vendor.Click();
  WorkspaceUtils.VPWSearchByValue(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  //WorkspaceUtils.VPWSearchByValue(vendor,"Vendor",VendorNo,"Vendor Number");
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

 
  var createselection = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McPlainCheckboxView.Button;
    createselection.Click();
    
  Log.Message(Duedate)
   var duedate = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McDatePickerWidget;
   Sys.HighlightObject(duedate);
   
//   if(duedate.getText()!=DueDate){
//      if(DueDate!=""){
       aqUtils.Delay(1000, Indicator.Text);
       duedate.setText(Duedate);
//          WorkspaceUtils.CalenderDateSelection(duedate,DueDate)
          ValidationUtils.verify(true,true,"Due Date is selected in Maconomy"); 
//        }
//    }
//    else{ 
//      ValidationUtils.verify(false,true,"Due Date is Needed to Create a Payment Selection");
//    } 
ReportUtils.logStep_Screenshot();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var scroll = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  scroll.MouseWheel(-200);
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//   if(ExchangeDate!=""){
//   var exchange = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite.McDatePickerWidget;
//   Sys.HighlightObject(exchange);
//      exchange.HoverMouse();
//      WorkspaceUtils.CalenderDateSelection(exchange,ExchangeDate)
//      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
//    }
//    else{ 
//      ValidationUtils.verify(false,true,"Exchange Date is Needed to Create a Payment Selection");
//    } 
    
    
//  var showall = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget4.Composite.McPlainCheckboxView.Button;
//  showall.Click();

  var layoutOPtion = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget4.Composite2.McPopupPickerWidget;
  layoutOPtion.Keys(layout);
//layout.Keys("Standard");
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(save)
  save.Click();
ReportUtils.logStep_Screenshot();
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var print = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  print.Click();
aqUtils.Delay(5000, Indicator.Text);
WorkspaceUtils.savePDF_And_WriteToExcel("PaymentSelectionMpl","PaymentSelection");
}


function validateCreateChangePaymentSelection_standardLayout()
{
  var fileName = "C:\\Users\\516188\\Documents\\Standard\\1008_CPS standed.pdf";
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
  var workBook = "C:\\Users\\516188\\Documents\\Standard\\DS_SPN_REGRESSION.xlsx";
  var country = "Spain";
  EnvParams.Opco = "1008";
 
  var pdflineSplit = docObj.split("\r\n");
 
  ExcelUtils.setExcelName(workBook, "CreatePaymentSelection", true);
  var vendorNumber = ReadExcelSheet("Vendor Number",EnvParams.Opco,"CreatePaymentSelection");
  var paymentAgent  = ReadExcelSheet("Payment_Agent",EnvParams.Opco,"CreatePaymentSelection");
  var paymodeMode = ReadExcelSheet("Paymode_Mode",EnvParams.Opco,"CreatePaymentSelection");
  var exchangeDate = ReadExcelSheet("ExchangeRateDate",EnvParams.Opco,"CreatePaymentSelection");
  var dueDate = ReadExcelSheet("Latest Due Date",EnvParams.Opco,"CreatePaymentSelection");
  var amount= ReadExcelSheet("Amount",EnvParams.Opco,"CreatePaymentSelection");
                    
  verifyVendorNumber(vendorNumber, pdflineSplit);     
  verifyPaymentAgent(paymentAgent, pdflineSplit);    
  verifyPaymodeMode(paymodeMode,pdflineSplit);          
  verifyExchangeDate(exchangeDate,pdflineSplit);
  verifyDueDate(dueDate,pdflineSplit);     
  verifyAmount(amount,pdflineSplit);
 }


function validateCreateChangePaymentSelection_wppLayout(filepathforMplValidation,workBook,sheetName)
{
  var fileName = filepathforMplValidation;
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
 // var workBook = "C:\\GlobalTestSuiteAutomation_Bank\\WppRegpack\\TestResource\\Regression\\DS_SPN_REGRESSION.xlsx";
 //  var country = "Spain";
  //EnvParams.Opco = "1006";
 
  var pdflineSplit = docObj.split("\r\n");
 
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var vendorNumber = ReadExcelSheet("Vendor Number",EnvParams.Opco,sheetName);
  var vendorInvoiceNo = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,sheetName);
  var amount= ReadExcelSheet("Amount",EnvParams.Opco,sheetName);
  var exchangeDate = ReadExcelSheet("ExchangeDate",EnvParams.Opco,sheetName);
  var dueDate = ReadExcelSheet("Due Date",EnvParams.Opco,sheetName);
  var paymodeMode = ReadExcelSheet("Paymode_Mode",EnvParams.Opco,sheetName);
               
                
  verifyVendorNumber(vendorNumber, pdflineSplit);
  verifyInvoiceNumber(vendorInvoiceNo,pdflineSplit);          
  verifyAmount(amount,pdflineSplit);
  verifyExchangeDate(exchangeDate,pdflineSplit);
  verifyDueDate(dueDate,pdflineSplit);     
  verifyPaymodeMode(paymodeMode,pdflineSplit);      
          
}



//Go To Job from Menu
function goToJobMenuItem(){

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.Banking.Exists()){
 ImageRepository.ImageSet.Banking.Click();// GL
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



function changePaymentSelection()
{
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
var ApprveTab =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.ApproveTab;
waitForObj(ApprveTab);

ApprveTab.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var PayToVendorFrom =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.PaytoVendorFromDate;
  Sys.HighlightObject(PayToVendorFrom);
  
  if(VendorNo!=""){
  PayToVendorFrom.Click();
//  WorkspaceUtils.VPWSearchByValue(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  WorkspaceUtils.VPWSearchByValue(PayToVendorFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
  }


var PayToVendorToValue  =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.PayToVendorTo

 if(VendorNo!=""){
  PayToVendorToValue.Click();
//  WorkspaceUtils.VPWSearchByValue(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  WorkspaceUtils.VPWSearchByValue(PayToVendorToValue,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
  }
  
  
    
  var company = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.CompanyFrom
  //Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
    waitForObj(company);
  Sys.HighlightObject(company)
  company.Click();
  WorkspaceUtils.SearchByValue(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

  aqUtils.Delay(1000, Indicator.Text);
  var company1 =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.CompanyTo
  // Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget2;
  Sys.HighlightObject(company1)
  company1.Click();
  WorkspaceUtils.SearchByValue(company1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

//    var duedate = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McDatePickerWidget;
//  duedate.Click();
// duedate.setText(" ");
//    
//     var duedate1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;  
//     duedate1.Click();
//     duedate1.setText(" ");
//     
   var paymentAgent = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;   
  paymentAgent.Click();
  paymentAgent.setText(" ");
//  
  var paymentMode = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;   
  paymentMode.Click(); 
  WorkspaceUtils.SearchByValue(paymentMode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Mode").OleValue.toString().trim(),Paymodemode,"Payment Mode")
//  paymentMode.setText(" ");
  
 // var latestDueDate =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.LatestDueDate;
  
  var DueDateFrom =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.DueDateFrom;
  DueDateFrom.setText(OldDuedate)
  var DueDateTo =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.DueDateTo;
  DueDateTo.setText(OldDuedate)
  
  
  aqUtils.Delay(1000, Indicator.Text);
//  var scroll =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
var scroll = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10
  scroll.MouseWheel(-1);
//  scroll.MouseWheel(-200);
  aqUtils.Delay(3000, Indicator.Text);
//  

if(ImageRepository.ImageSet_Banking.EntriesDown.Exists()){
 ImageRepository.ImageSet_Banking.EntriesDown.Click();// GL
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
  var showEntriesCheckBox =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPlainCheckboxView.ShowEntriesCheckBox;
  
    if(!showEntriesCheckBox.getSelection()){ 
  showEntriesCheckBox.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showEntriesCheckBox.Click();
  ReportUtils.logStep("INFO", "showEntriesCheckBoxis UnChecked");
    Log.Message("showEntriesCheckBox is UnChecked")
    checkmark = true;
  }
  
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
} 
  var EntriesUp =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.EntriesUpButton;
  aqUtils.Delay(1000, Indicator.Text);
  EntriesUp.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var EntriesTable = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.EntriesTable;
waitForObj(EntriesTable);
       Sys.HighlightObject(EntriesTable);       
    
        var  column = EntriesTable.getColumnCount();
        var row = EntriesTable.getItemCount()
        Log.Message(column)
        Log.Message(row)
        
        var flag = false;
       Log.Message(InvoiceNo)
        for(var i=0;i<row;i++){
          Log.Message(EntriesTable.getItem(i).getText(2).OleValue.toString().trim())
          if(EntriesTable.getItem(i).getText(2).OleValue.toString().trim()==InvoiceNo){
             flag = true
            Log.Message(EntriesTable.getItem(i).getText(2).OleValue.toString().trim())
            ValidationUtils.verify(true,true,"Invoice Number is available in the table");
            EntriesTable.Keys("[Tab][Tab][Tab]");
              if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
              }
            break;
          }
          else{
            EntriesTable.Keys("[Down]");
          }
        }       
             
       
       
    aqUtils.Delay(1000, Indicator.Text);
       
       Log.Message(flag)
  if(flag){
  ValidationUtils.verify(flag,true,"Invoice No is Present in Table");
  TextUtils.writeLog("Invoice No is Present in Table");
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//     var DueDateTable = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.EntriesTable.DueDateTableField
var Date = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
Date.Click();
      Date.setText(Duedate); 
//       DueDateTable.setText(DueDate);
       
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
       var savebutton =Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SaveButton;
       
       savebutton.Click();
       
        aqUtils.Delay(5000, Indicator.Text);
          if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
          }
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management",Duedate)
       
}

}

function test()
{
//  var DueDate = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.EntriesTable.DueDateTableField
//       
//       DueDate.setText(" 1/12/12");
       
       
       var EntriesTable = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.EntriesTable;
       EntriesTable.Keys("[Tab][Tab][Tab]");
      EntriesTable.Keys(" 1/12/12");
}


function goToAPMenuItem(){

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountPayable.Exists()){
 ImageRepository.ImageSet.AccountPayable.Click();// GL
}
else if(ImageRepository.ImageSet.AccountPayable1.Exists()){
ImageRepository.ImageSet.AccountPayable1.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
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
Client_Managt.ClickItem("|AP Lookups");
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Lookups").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Lookups").OleValue.toString().trim());
Client_Managt.DblClickItem("|AP Lookups");
}
}
ReportUtils.logStep("INFO", "Moved to Banking Transactions from job Menu");
TextUtils.writeLog("Entering into Banking Transactions from Jobs Menu");
}



//Main Function
function ChangePayment_Main() {
TextUtils.writeLog("Create Payment Selection Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

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
sheetName = "ChangePaymentSelection";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
VendorNo,Paymentagent,Paymodemode,InvoiceNo,layout,DueDate ="";
OldDuedate = "";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

//try{
getDetails();
//goToAPMenuItem();
//goToAp();
//closeAllWorkspaces();
goToJobMenuItem();   
changePaymentSelection();
closeAllWorkspaces();
goToJobMenuItem();
CreatePaymentSeletion(); 
closeAllWorkspaces();
//}
//  catch(err){
//    Log.Message(err);
//  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

function verifyVendorNumber(vendorNumber,pdflineSplit)
{
    var vendorNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(vendorNumber))
             {
             Log.Message(vendorNumber+" vendorNumber is matching with Pdf");
             vendorNoFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !vendorNoFound)
          ValidationUtils.verify(false,true,"VendorNumber is not same in CreatePaymentFile");
  }  
}

function verifyInvoiceNumber(vendorInvoiceNo,pdflineSplit)
{
  var vendorInvoiceNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
          if(vendorInvoiceNo.includes(pdflineSplit[j]))             {
             Log.Message(vendorInvoiceNo+" vendorInvoiceNo is matching with Pdf");
             vendorInvoiceNoFound = true;
             break;
             }
         else
         continue;
         if(j==pdflineSplit.length-1 && !vendorInvoiceNoFound)
          ValidationUtils.verify(false,true,"vendorInvoiceNo is not same in CreatePaymentFile");
    
  }       
}

function verifyAmount(amount,pdflineSplit)
{
  var amountFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(amount))
             {
             Log.Message(amount+" amount is matching with Pdf");
             amountFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !amountFound)
          ValidationUtils.verify(false,true,"amount is not same in CreatePaymentFile");
    
    }
}

function verifyExchangeDate(exchangeDate,pdflineSplit)
{
  var exchangeDateFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(exchangeDate))
             {
             Log.Message(exchangeDate+" exchangeDate is matching with Pdf");
             exchangeDateFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !exchangeDateFound)
          ValidationUtils.verify(false,true,"exchangeDate is not same in CreatePaymentFile");
    
    } 
}

function verifyDueDate(dueDate,pdflineSplit)
{
     var dueDateFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(dueDate))
             {
             Log.Message(dueDate+" DueDate is matching with Pdf");
             dueDateFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !dueDateFound)
          ValidationUtils.verify(false,true,"DueDate is not same in CreatePaymentFile");
    
    }    
}
function verifyPaymentNumber(paymentNumber,pdflineSplit)
{
   var paymentNumberFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(paymentNumber))
             {
             Log.Message(paymentNumber+" PaymentNumber is matching with Pdf");
             paymentNumberFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymentNumberFound)
          ValidationUtils.verify(false,true,"PaymentNumber is not same in PrintReimmittance");    
    }   
}

function verifyPaymodeMode(paymodeMode, pdflineSplit)
{
   var paymodeModeFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(paymodeMode))
             {
             Log.Message(paymodeMode+" paymodeMode is matching with Pdf");
             paymodeModeFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymodeModeFound)
          ValidationUtils.verify(false,true,"paymodeMode is not same in CreatePaymentSelection/ChangePaymentSelection");    
    }
}
function verifyPaymentAgent(paymentAgent,pdflineSplit)
{
   var paymentAgentFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(paymentAgent))
             {
             Log.Message(paymentAgent+" paymentAgent is matching with Pdf");
             paymentAgentFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymentAgentFound)
          ValidationUtils.verify(false,true,"paymentAgent is not same in CreatePaymentSelection/ChangePaymentSelection");    
    }
}



