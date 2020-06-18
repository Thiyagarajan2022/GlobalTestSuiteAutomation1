//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreatePaymentFile";
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
var Paymentmode="";



//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
Paymentagent = ExcelUtils.getRowDatas("Payment_Agent",EnvParams.Opco)
Log.Message(Paymentagent)
if((Paymentagent==null)||(Paymentagent=="")){ 
ValidationUtils.verify(false,true,"Payment Agent is Needed to Create a Payment File");
}
Paymentmode = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco)
Log.Message(Paymentmode)
if((Paymentmode==null)||(Paymentmode=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment File");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
Log.Message(VendorNo)
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment File");
}
Duedate = ReadExcelSheet("DueDate",EnvParams.Opco,"Data Management");
Log.Message(Duedate)
if((Duedate=="")||(Duedate==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Duedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
}
if((Duedate==null)||(Duedate=="")){ 
ValidationUtils.verify(false,true,"DueDate is Needed to Create a Payment File");
}
}




function PaymentFile() {
  ReportUtils.logStep("INFO", "Enter Payment File Details");
 var banking = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(banking);
  WorkspaceUtils.waitForObj(banking);
  var approve = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;
  approve.HoverMouse();
  approve.Click();
  ReportUtils.logStep_Screenshot("");
  var vendor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  Sys.HighlightObject(vendor);
  if(VendorNo!=""){
  vendor.Click();
  WorkspaceUtils.VPWSearchByValue(vendor,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment File");
  }
  
  var vendor1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget2;
  Sys.HighlightObject(vendor1);
  if(VendorNo!=""){
  vendor1.Click();
  WorkspaceUtils.VPWSearchByValue(vendor1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  TextUtils.writeLog("Vendor is selected from macanomy:"+VendorNo+"");  
  }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment File");
  }
  
  
  var company = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;    waitForObj(company);
  Sys.HighlightObject(company)
  company.Click();
  WorkspaceUtils.SearchByValue(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

  aqUtils.Delay(1000, Indicator.Text);
  
  var company1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget2;  Sys.HighlightObject(company1)
  company1.Click();
  WorkspaceUtils.SearchByValue(company1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
  
  Log.Message(Duedate)
  var duedate = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McDatePickerWidget;
  duedate.Click();
  if((Duedate!="")&&(Duedate!=null)){
       aqUtils.Delay(1000, Indicator.Text);
          WorkspaceUtils.CalenderDateSelection(duedate,Duedate)
          ValidationUtils.verify(true,true,"Due Date is selected in Maconomy"); 
        }
    else{ 
      ValidationUtils.verify(false,true,"Payment Date is Needed  for Remittance Email");
    } 
    
     var duedate1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;  duedate1.Click();
   if((Duedate!="")&&(Duedate!=null)){
       aqUtils.Delay(1000, Indicator.Text);
          WorkspaceUtils.CalenderDateSelection(duedate1,Duedate)
          ValidationUtils.verify(true,true,"Due Date is selected in Maconomy"); 
        }
    else{ 
      ValidationUtils.verify(false,true,"Payment Date is Needed  for Remittance Email");
    } 
  
  var paymentAgent = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;   
  if(Paymentagent!=""){
  paymentAgent.Click();
  WorkspaceUtils.SearchByValue(paymentAgent,"Payment Agent",Paymentagent,"Payment Agent")
  }
  else{ 
    ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment File");
  }
  
  var paymentMode = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;   
  if(Paymentmode!=""){
  paymentMode.Click();
  WorkspaceUtils.SearchByValue(paymentMode,"Payment Mode",Paymentmode,"Payment Mode")
  }
  else{ 
    ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment File");
  }

 
  var scroll = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10;  scroll.MouseWheel(-200);
  scroll.MouseWheel(-200);
  aqUtils.Delay(1000, Indicator.Text);
  
  var show1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPlainCheckboxView.Button;
   if(show1.getSelection()){ 
  show1.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  show1.Click();
  ReportUtils.logStep("INFO", "Show Only Entries without Payer Id.");
    Log.Message("Show Only Entries without Payer Id.")
    checkmark = true;
  }
  
  var show2 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McPlainCheckboxView.Button;
   if(show2.getSelection()){ 
  show2.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  show2.Click();
  ReportUtils.logStep("INFO", "Show Only Entries without Card Type Code");
    Log.Message("Show Only Entries without Card Type Code")
    checkmark = true;
  }
  
  var show3 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPlainCheckboxView.Button;
  if(show3.getSelection()){ 
  show3.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  show3.Click();
  ReportUtils.logStep("INFO", "Show Only Entries Selected for Payment");
    Log.Message("Show Only Entries Selected for Payment")
    checkmark = true;
  }
  
  var donot = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McPlainCheckboxView.Button;
  if(donot.getSelection()){ 
  donot.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  donot.Click();
  ReportUtils.logStep("INFO", "Do Not Show Entries Being Paid");
    Log.Message("Do Not Show Entries Being Paid")
    checkmark = true;
  }
  
  var showentry = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McPlainCheckboxView.Button;
  if(showentry.getSelection()){ 
  showentry.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  ReportUtils.logStep("INFO", "Show Entries");
    Log.Message("Show Entries")
    checkmark = true;
  }
  else{
    showentry.Click();
  }
   
  
  var save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl;  Sys.HighlightObject(save)
  save.Click();
  
  var entries = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(entries)
  
  var ApproveConPayment = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl2;
  Sys.HighlightObject(ApproveConPayment);
  ApproveConPayment.Click();
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
Client_Managt.ClickItem("|Bank Transactions");
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions").OleValue.toString().trim());
Client_Managt.DblClickItem("|Bank Transactions");
}
}
ReportUtils.logStep("INFO", "Moved to Banking Transactions from job Menu");
TextUtils.writeLog("Entering into Banking Transactions from Jobs Menu");
}




//Main Function
function CreatePaymentFile() {
TextUtils.writeLog("Create Payment File Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior APs","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreatePaymentFile";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
VendorNo,Paymentagent,Paymentmode,DueDate ="";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

  try{
    getDetails();
    goToJobMenuItem();   
    PaymentFile(); 
  }
    catch(err){
      Log.Message(err);
    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



