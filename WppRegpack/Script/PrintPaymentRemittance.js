//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "PrintPaymentRemittance";
var Language = "";
Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var DueDate="";
var VendorNo="";
var Paymentagent="";
var Paymodemode="";
var ExchangeDate="";
var layoutTypes="";
var Invoicenumber="";
var PaymentDate="";
var PaymentNo="";
var filepathforMplValidation ="";
var amount = ""
//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
Log.Message(sheetName)

ExcelUtils.setExcelName(workBook, "Data Management", true);
Paymentagent = ReadExcelSheet("Payment Agent",EnvParams.Opco,"Data Management");
if((Paymentagent=="")||(Paymentagent==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Paymentagent = ExcelUtils.getRowDatas("Payment_Agent",EnvParams.Opco)
}
Log.Message(Paymentagent)
if((Paymentagent==null)||(Paymentagent=="")){ 
ValidationUtils.verify(false,true,"Payment Agent is Needed to Create a Payment Selection");
}

//Paymodemode = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco)
//Log.Message(Paymodemode)
//if((Paymodemode==null)||(Paymodemode=="")){ 
//ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment Selection");
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

//PaymentDate = ExcelUtils.getRowDatas("Payment Date",EnvParams.Opco)
//Log.Message(PaymentDate)
//if((PaymentDate==null)||(PaymentDate=="")){ 
//ValidationUtils.verify(false,true,"PaymentDate is Needed to Create a Payment Selection");
//}

ExcelUtils.setExcelName(workBook, "Data Management", true);
PaymentDate = ReadExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management");
if((PaymentDate=="")||(PaymentDate==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
PaymentDate = ExcelUtils.getRowDatas("Payment Date",EnvParams.Opco)
}
Log.Message(PaymentDate)
if((PaymentDate==null)||(PaymentDate=="")){ 
ValidationUtils.verify(false,true,"Payment Date is Needed to Create a Payment Selection");
}


//PaymentNo = ExcelUtils.getRowDatas("Payment_Number",EnvParams.Opco)
//Log.Message(PaymentNo)
//if((PaymentNo==null)||(PaymentNo=="")){ 
//ValidationUtils.verify(false,true,"PaymentNumber is Needed to Create a Payment Selection");
//}

ExcelUtils.setExcelName(workBook, sheetName, true);
PrintLayout = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(PrintLayout)
if((PrintLayout==null)||(PrintLayout=="")){ 
ValidationUtils.verify(false,true,"PrintLayout is Needed to Create a Payment Selection");
}
//DueDate = ExcelUtils.getRowDatas("Due_Date",EnvParams.Opco)
//if((DueDate==null)||(DueDate=="")){ 
//ValidationUtils.verify(false,true,"Due Date Number is Needed to Create a Payment Selection");
//}
//ExchangeDate = ExcelUtils.getRowDatas("Exchange_Date",EnvParams.Opco)
//Log.Message(ExchangeDate)
//if((ExchangeDate==null)||(ExchangeDate=="")){ 
//ValidationUtils.verify(false,true,"Exchange Date  is Needed to Create a Payment Selection");
//}
//layoutTypes = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
//Log.Message(layoutTypes)
//if((layoutTypes==null)||(layoutTypes=="")){ 
//ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Selection");
//}


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
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
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


}


function printPaymentRemittance()
{
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var lookups =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.LookUps;
waitForObj(lookups);
lookups.Click();
ReportUtils.logStep_Screenshot("");
aqUtils.Delay(1000);
var paymentDateFrom =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.PaymentDateFrom;

paymentDateFrom.setText(PaymentDate);
aqUtils.Delay(1000);
var paymentDateTo =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.PaymentDateTo;

paymentDateTo.setText(PaymentDate);
 
 
var vendorNoFrom = Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.VendorNoFrom;

  if(VendorNo!=""){
  vendorNoFrom.Click();
  WorkspaceUtils.VPWSearchByValue(vendorNoFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
 // WorkspaceUtils.VPWSearchByValue(vendorNoFrom,"Vendor",VendorNo,"Vendor Number");
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
  }

var vendorNoTo = Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.VnedorNoTo;


  if(VendorNo!=""){
  vendorNoTo.Click();
  WorkspaceUtils.VPWSearchByValue(vendorNoTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
//  WorkspaceUtils.VPWSearchByValue(vendorNoTo,"Vendor",VendorNo,"Vendor Number");
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
  }
  
//var paymentNoFrom =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.PaymentNoFrom;
//
// if(PaymentNo!=""){
//  paymentNoFrom.Click();
//  WorkspaceUtils.SearchByValue(paymentNoFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Exported Files").OleValue.toString().trim(),PaymentNo,"Output Data Number");
//    }
// else{ 
//    ValidationUtils.verify(false,true,"PaymentNo is Needed to Create a Payment Selection");
//  }
//  
//var PaymentNoTo =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.PaymentNoTo;
//
// if(PaymentNo!=""){
//  PaymentNoTo.Click();
//  WorkspaceUtils.SearchByValue(PaymentNoTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Exported Files").OleValue.toString().trim(),PaymentNo,"Output Data Number");
//    }
// else{ 
//    ValidationUtils.verify(false,true,"PaymentNo is Needed to Create a Payment Selection");
//  }
//  
  
var showpaid =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.McPlainCheckboxView.ShowPaidCheckBox;
var showNonClosed =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite2.McPlainCheckboxView.ShowNonClosed;
var showPaymentSelection = Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite3.McPlainCheckboxView.ShowPaymentSelection;
var saveButton =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SaveButton;
var showErrorReportedReversed =Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite4.McPlainCheckboxView.ErrorReportedReveresed;


 if(!showpaid.getSelection()){ 
  showpaid.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showpaid.Click();
  ReportUtils.logStep("INFO", "showpaid is UnChecked");
    Log.Message("showpaid is UnChecked")
    checkmark = true;
  }
  
  
 if(!showNonClosed.getSelection()){ 
  showNonClosed.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showNonClosed.Click();
  ReportUtils.logStep("INFO", "showNonClosed is UnChecked");
    Log.Message("showNonClosed is UnChecked")
    checkmark = true;
  }
  
  
   if(showErrorReportedReversed.getSelection()){ 
  showErrorReportedReversed.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showErrorReportedReversed.Click();
  ReportUtils.logStep("INFO", "showErrorReportedReversed is UnChecked");
    Log.Message("showErrorReportedReversed is UnChecked")
    checkmark = true;
  }
  
   if(!showPaymentSelection.getSelection()){ 
  showPaymentSelection.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showPaymentSelection.Click();
  ReportUtils.logStep("INFO", "showPaymentSelection is UnChecked");
    Log.Message("showPaymentSelection is UnChecked")
    checkmark = true;
  }
  
  
aqUtils.Delay(1000);
saveButton.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  

}


PaymentDate = formatDate(PaymentDate);

var table  = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    
  Log.Message(PaymentDate)
  Log.Message(table.getItem(v).getText_2(0).OleValue.toString().trim()==PaymentDate) 
  Log.Message(amount)
  Log.Message(table.getItem(v).getText_2(5).OleValue.toString().trim()==amount)
  Log.Message(Paymodemode)
  Log.Message(table.getItem(v).getText_2(10).OleValue.toString().trim()==Paymodemode)
  
  if((table.getItem(v).getText_2(0).OleValue.toString().trim()==PaymentDate) && 
  (table.getItem(v).getText_2(1).OleValue.toString().trim()==VendorNo) && 
  (table.getItem(v).getText_2(10).OleValue.toString().trim()==Paymodemode)){ 
  PaymentNo = table.getItem(v).getText_2(4).OleValue.toString().trim()  
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }

  if(flag)
  {
    ExcelUtils.setExcelName(workBook,"Data Management", true);
    ExcelUtils.WriteExcelSheet("Payment Number",EnvParams.Opco,"Data Management",PaymentNo)

  }

var print = Aliases.Maconomy.PrintRemittance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.PrintPaymentRemittance
aqUtils.Delay(1000);
print.Click();

//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions - Payment Orders").OleValue.toString().trim())    
//{
//var button = Sys.Process("Maconomy").SWTObject("Shell", "Bank Transactions - Payment Orders").SWTObject("Composite", "", 2).SWTObject("Button",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Bank Transactions - Payment Orders").SWTObject("Label", "*").WndCaption;
//      Log.Message(label );
//       button.HoverMouse();
//     ReportUtils.logStep_Screenshot("");
//      button.Click();
//    aqUtils.Delay(1000);
//    
//}

var PrintPopup =Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Payment Order").OleValue.toString().trim());

waitForObj(PrintPopup);
aqUtils.Delay(15000);


var paymentDateFrom =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.PaymentDateFrom;
paymentDateFrom.setText(PaymentDate);

var PayemntDateTo =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.PaymentDateTo;
PayemntDateTo.setText(PaymentDate);

var PaymentNoFrom =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.PaymentNoFrom
PaymentNoFrom.Click();
PaymentNoFrom.setText(PaymentNo);

var PaymentNoTo =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.PaymentNoTo;
PaymentNoTo.Click();
PaymentNoTo.setText(PaymentNo);

var scroll1 =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10;

var includePaid =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McPlainCheckboxView.IncludePaidCheckbox;
var errorReportedfailed =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPlainCheckboxView.IncludeNonClosed;
var includeNonClosed =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McPlainCheckboxView.NonClosed;
var includePaymentSelection =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPlainCheckboxView.IncludePaymentSelection;

  if(!includePaid.getSelection()){ 
  includePaid.HoverMouse();
ReportUtils.logStep_Screenshot("");
  includePaid.Click();
  ReportUtils.logStep("INFO", "includePaid is UnChecked");
    Log.Message("includePaid is UnChecked")
    checkmark = true;
  }
  
    if(errorReportedfailed.getSelection()){ 
  errorReportedfailed.HoverMouse();
ReportUtils.logStep_Screenshot("");
  errorReportedfailed.Click();
  ReportUtils.logStep("INFO", "errorReportedfailed is UnChecked");
    Log.Message("errorReportedfailed is UnChecked")
    checkmark = true;
  }
  
      if(!includeNonClosed.getSelection()){ 
  includeNonClosed.HoverMouse();
ReportUtils.logStep_Screenshot("");
  includeNonClosed.Click();
  ReportUtils.logStep("INFO", "includeNonClosed is UnChecked");
    Log.Message("includeNonClosed is UnChecked")
    checkmark = true;
  }

  
      if(!includePaymentSelection.getSelection()){ 
  includePaymentSelection.HoverMouse();
ReportUtils.logStep_Screenshot("");
  includePaymentSelection.Click();
  ReportUtils.logStep("INFO", "includePaymentSelection is UnChecked");
    Log.Message("includePaymentSelection is UnChecked")
    checkmark = true;
  }

  scroll1.MouseWheel(-200);
  aqUtils.Delay(1000, Indicator.Text);


var layout =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.LayoutOption;

layout.setText(PrintLayout);


var pagebreakCheckBox =Aliases.Maconomy.PrintRemittancePopup.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McPlainCheckboxView.PageBreakCheckBox;

     if(!pagebreakCheckBox.getSelection()){ 
  pagebreakCheckBox.HoverMouse();
ReportUtils.logStep_Screenshot("");
  pagebreakCheckBox.Click();
  ReportUtils.logStep("INFO", "pagebreakCheckBox is UnChecked");
    Log.Message("pagebreakCheckBox is UnChecked")
    checkmark = true;
  }

var PrintPDF = Aliases.Maconomy.PrintRemittancePopup.Composite.Composite2.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print").OleValue.toString().trim())
//Aliases.Maconomy.PrintRemittancePopup.Composite.Composite2.PrintButton;

PrintPDF.Click();


var SaveTitle = "";
var sFolder = "";
//NameMapping.Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PaymentOrder-3.pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5)
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PaymentOrder"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PaymentOrder"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("P_PaymentOrder")!=-1){
    aqUtils.Delay(2000, Indicator.Text);

Sys.HighlightObject(pdf)
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
    
if(ImageRepository.PDF.ChooseFolder.Exists())
ImageRepository.PDF.ChooseFolder.Click();
else{ 
var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
WorkspaceUtils.waitForObj(window);

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x73); //F4
Sys.Desktop.KeyUp(0x12); //Alt
Sys.Desktop.KeyUp(0x73); //F4
aqUtils.Delay(2000, Indicator.Text);
Sys.HighlightObject(pdf);

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
}
var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
aqUtils.Delay(2000, Indicator.Text);
SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");

filepathforMplValidation =sFolder+SaveTitle+".pdf";
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print Remittance PDF is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PrintPaymentRemittanceMpl",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")

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
  company.Keys("[Tab][Tab][Tab][Tab][Tab]");
//  var vendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var invoice = Aliases.Maconomy.AR.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;  
//Aliases.Maconomy.AR.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(invoice)
  invoice.setText(Invoicenumber);
 
        for(var i=0;i<table.getItemCount();i++){          
          if(table.getItem(i).getText_2(5).OleValue.toString().trim()==Invoicenumber){
            break;
          }  
          else{
              table.Keys("[Down]");
          } 
        } 
         
        ReportUtils.logStep_Screenshot();    
        
  aqUtils.Delay(5000,Indicator.Text);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x46);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x46);  
   
  var Duedate1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  Sys.HighlightObject(Duedate1)  
//  waitForObj(Duedate1);
  var Duedatee = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget2;
  waitForObj(Duedatee)
    aqUtils.Delay(2000,"Waiting for window");
  Sys.HighlightObject(Duedatee)
  DueDate = Duedatee.getText();
  Log.Message(DueDate)
  ReportUtils.logStep_Screenshot();
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//  ExcelUtils.WriteExcelSheet("DueDate",EnvParams.Opco,"Data Management",Duedatee)
  
}



function CreatePaymentSeletion() {
ReportUtils.logStep("INFO", "Enter Bank Details");
 var banking = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(banking);
  WorkspaceUtils.waitForObj(banking);
  var create = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  create.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  var vendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  Sys.HighlightObject(vendor);
  if(VendorNo!=""){
  vendor.Click();
//  WorkspaceUtils.VPWSearchByValue(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  WorkspaceUtils.VPWSearchByValue(vendor,"Vendor",VendorNo,"Vendor Number");
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

  
  var paymentAgent = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
   if(Paymentagent!=""){
  paymentAgent.Click();
  WorkspaceUtils.SearchByValue(paymentAgent,"Payment Agent",Paymentagent,"Payment Agent")
}else{ 
  ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment Selection");
}
  
  Log.Message(Paymodemode)
  var paymentMode = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget;
   if(Paymodemode!=""){
  paymentMode.Click();
  WorkspaceUtils.SearchByValue(paymentMode,"Payment Mode",Paymodemode,"Payment Mode")
  }else{ 
    ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment Selection");
  }

 
  var createselection = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McPlainCheckboxView.Button;
    createselection.Click();
    
  Log.Message(DueDate)
   var duedate = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McDatePickerWidget;
   Sys.HighlightObject(duedate);
   
   if(duedate.getText()!=DueDate){
      if(DueDate!=""){
       aqUtils.Delay(1000, Indicator.Text);
       duedate.setText(DueDate);
//          WorkspaceUtils.CalenderDateSelection(duedate,DueDate)
          ValidationUtils.verify(true,true,"Due Date is selected in Maconomy"); 
        }
    }
    else{ 
      ValidationUtils.verify(false,true,"Due Date is Needed to Create a Payment Selection");
    } 
ReportUtils.logStep_Screenshot();
 var scroll = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  scroll.MouseWheel(-200);
  aqUtils.Delay(1000, Indicator.Text);
  
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

  var layout = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget4.Composite2.McPopupPickerWidget;
  layout.setText(layoutTypes);
//  layout.Keys("WPP Payment");
//layout.Keys("Standard");
  
  var save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(save)
  save.Click();
ReportUtils.logStep_Screenshot();
  var print = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  print.Click();

var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "PaymentSelection"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "PaymentSelection"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("PaymentSelection")!=-1){
    aqUtils.Delay(2000, Indicator.Text);

Sys.HighlightObject(pdf)
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
    
if(ImageRepository.PDF.ChooseFolder.Exists())
ImageRepository.PDF.ChooseFolder.Click();
else{ 
var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
WorkspaceUtils.waitForObj(window);

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x73); //F4
Sys.Desktop.KeyUp(0x12); //Alt
Sys.Desktop.KeyUp(0x73); //F4
aqUtils.Delay(2000, Indicator.Text);
Sys.HighlightObject(pdf)

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
}
var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
aqUtils.Delay(2000, Indicator.Text);
SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");

var filepathforMplValidation =sFolder+SaveTitle+".pdf";
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
saveAs.Click();
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print Draft Quote is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

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
function PrintPaymentRemittance_Main() {
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
sheetName = "PrintPaymentRemittance";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
VendorNo,Paymentagent,Paymodemode,ExchangeDate,layoutTypes,Invoicenumber,PaymentDate,PaymentNo,filepathforMplValidation,amount ="";

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
printPaymentRemittance();
closeAllWorkspaces();
//CreatePaymentSeletion(); 
//}
//  catch(err){
 //   Log.Message(err);
  //}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}





function formatDate(d){ 
  var dd=""
if(d.includes("/"))
{
  var parts = d.split("/");
  if(parts[0].indexOf("0")==0){
   parts[0] = parts[0].replace("0", "");
  }
  if(parts[1].indexOf("0")==0){
    parts[1] = parts[1].replace("0", "");
  }
  dd = parts[0]+"/"+parts[1]+"/"+parts[2];
  Log.Message(dd)
  }
  if(d.includes("-"))
  {
  var parts = d.split("-");
  if(parts[0].indexOf("0")==0){
   parts[0] = parts[0].replace("0", "");
  }
  if(parts[1].indexOf("0")==0){
    parts[1] = parts[1].replace("0", "");
  }
  dd = parts[0]+"-"+parts[1]+"-"+parts[2];
  Log.Message(dd)
  }
  return dd;
}

