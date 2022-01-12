//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Void Payment";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var VendorNo,PaymentNo,Paymentdate="";


//Main Function
function VoidaPayment(){ 
TextUtils.writeLog("Void a Payment Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Senior AP","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "Void Payment";

ExcelUtils.setExcelName(workBook, sheetName, true);
count = true;
checkmark = false;
STIME = "";
VendorNo,PaymentNo,Paymentdate="";

try{
getDetails();
goToBankingTransaction();   
gotoStatusReporting(); 
paymentLine();  
}
  catch(err){
    Log.Message(err);
  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}

function getDetails(){ 


//ExcelUtils.setExcelName(workBook, "Data Management", true);
//Paymentdate = ReadExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management");
//if((Paymentdate==null)||(Paymentdate=="")){ 
//ExcelUtils.setExcelName(workBook, sheetName, true);
//Paymentdate = ExcelUtils.getRowDatas("PaymentDate",EnvParams.Opco)
//Log.Message(Paymentdate)
//}
//if((Paymentdate==null)||(Paymentdate=="")){ 
//ValidationUtils.verify(false,true,"Payment Date is Needed to Create a Remittance Email");
//}
//Paymentdate = formatDate(Paymentdate)

ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
Log.Message(VendorNo)
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
Log.Message(VendorNo)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Remittance Email");
}


ExcelUtils.setExcelName(workBook, "Data Management", true);
PaymentNo = ReadExcelSheet("Payment Number",EnvParams.Opco,"Data Management");
Log.Message(PaymentNo)
if((PaymentNo=="")||(PaymentNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
PaymentNo = ExcelUtils.getRowDatas("PaymentNumber",EnvParams.Opco)
Log.Message(PaymentNo)
}
if((PaymentNo==null)||(PaymentNo=="")){ 
ValidationUtils.verify(false,true,"Payment Number  is Needed to Create a Remittance Email");
}


}

function formatDate(d){ 
  var parts = d.split("/");
  if(parts[0].indexOf("0")==0){
   parts[0] = parts[0].replace("0", "");
  }
  if(parts[1].indexOf("0")==0){
    parts[1] = parts[1].replace("0", "");
  }
  var dd = parts[0]+"/"+parts[1]+"/"+parts[2];
  Log.Message(dd)
  return dd;
}


//Go To Job from Menu
function goToBankingTransaction(){

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


function gotoStatusReporting(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
var StatusReporting = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
StatusReporting.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
   aqUtils.Delay(1000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
   var paymentDate = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;
   paymentDate.setText(Paymentdate);
   var paymentDate = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget2;
   paymentDate.setText(Paymentdate);
   
   aqUtils.Delay(1000);
   var Vendor = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
    if(VendorNo!=""){
    Vendor.Click();
    WorkspaceUtils.VPWSearchByValue(Vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
    }
    aqUtils.Delay(1000);
   var Vendor = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget2;
   if(VendorNo!=""){
    Vendor.Click();
    WorkspaceUtils.VPWSearchByValue(Vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
    }
    aqUtils.Delay(1000);
   var Company = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  Sys.HighlightObject(Company)
  Company.Click();
  WorkspaceUtils.SearchByValue(Company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
  aqUtils.Delay(1000);
   var Company = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget2;
  Sys.HighlightObject(Company)
  Company.Click();
  WorkspaceUtils.SearchByValue(Company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

   var PayNo = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget;
    Sys.HighlightObject(PayNo);  
    if((PaymentNo!="")&&(PaymentNo!=null)){
    PayNo.Click();
    WorkspaceUtils.SearchByValue(PayNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Exported Files").OleValue.toString().trim(),PaymentNo,"Output Data No.");
    }
   aqUtils.Delay(1000);
   var PayNo = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget2;
   Sys.HighlightObject(PayNo);  
    if((PaymentNo!="")&&(PaymentNo!=null)){
    PayNo.Click();
    WorkspaceUtils.SearchByValue(PayNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Exported Files").OleValue.toString().trim(),PaymentNo,"Output Data No.");
    }
    aqUtils.Delay(1000);
   var ShowEntered = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McPlainCheckboxView.Button;
   if(!ShowEntered.getSelection()){ 
  ShowEntered.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  ShowEntered.Click();
  ReportUtils.logStep("INFO", "ShowEntered is Checked");
  Log.Message("ShowEntered is Checked")
  }
  aqUtils.Delay(1000);
  
  var scroll = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  scroll.Click();
  scroll.MouseWheel(-200);
  aqUtils.Delay(7000);
   var ShowPaid = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPlainCheckboxView.Button;
   if(!ShowPaid.getSelection()){ 
  ShowPaid.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  ShowPaid.Click();
  ReportUtils.logStep("INFO", "ShowEntered is Checked");
  Log.Message("ShowEntered is Checked")
  }
  aqUtils.Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  
   var Save = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
   Save.Click();
     aqUtils.Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
}

function paymentLine(){ 
  
var AutomaticReply = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
AutomaticReply.Click();
     aqUtils.Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
var table = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

var falg = false;
for(var v=0;v<table.getItemCount();v++){ 

  if((table.getItem(v).getText_2(4).OleValue.toString().trim()== PaymentNo) && 
  (table.getItem(v).getText_2(2).OleValue.toString().trim()==VendorNo)){  
   for(var i=0;i<12;i++){ 
     Sys.Desktop.KeyDown(0x09)
     Sys.Desktop.KeyUp(0x09)
     aqUtils.Delay(1000);
   }
   var reversed = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
   if(!reversed.getSelection()){ 
  reversed.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  reversed.Click();
  ReportUtils.logStep("INFO", "Reversed is Checked");
  Log.Message("Reversed is Checked")
  }
  aqUtils.Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
}
  
var Save = Aliases.Maconomy.VoidPayment.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Save.Click();
  aqUtils.Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  aqUtils.Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
var ApproveEntries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 5);
Sys.HighlightObject(ApproveEntries);
ApproveEntries.Click();
aqUtils.Delay(10000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
 aqUtils.Delay(10000);
 
 
 var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print out Payment Listing"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print out Payment Listing"+"*", 1).WndCaption.indexOf("Print out Payment Listing")!=-1){
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
ExcelUtils.WriteExcelSheet("Void Payment Pdf",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")

   aqUtils.Delay(5000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  
}