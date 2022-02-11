//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "EmailRemittance";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var VendorNo,PaymentNo,Paymentdate="";


//Main Function
function RemittanceEmail() {
TextUtils.writeLog("Create Remittance Email Started"); 
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
sheetName = "EmailRemittance";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
VendorNo,PaymentNo,Paymentdate="";

try{
getDetails();
goToJobMenuItem();   
Remittance();   
}
  catch(err){
    Log.Message(err);
  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}




//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);

//Paymentdate = ExcelUtils.getRowDatas("PaymentDate",EnvParams.Opco)
//Log.Message(Paymentdate)
//if((Paymentdate==null)||(Paymentdate=="")){ 
//ValidationUtils.verify(false,true,"Payment Date is Needed to Create a Remittance Email");
//}

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





function Remittance() {
  ReportUtils.logStep("INFO", "Enter Remittance Email Details");
  var Email = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl; Sys.HighlightObject(Email);
  WorkspaceUtils.waitForObj(Email);
  Sys.HighlightObject(Email);
  
  var paymentnumber = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  Sys.HighlightObject(paymentnumber);  
    if((PaymentNo!="")&&(PaymentNo!=null)){
  paymentnumber.Click();
    WorkspaceUtils.SearchByValue(paymentnumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Exported Files").OleValue.toString().trim(),PaymentNo,"Output Data No.");
  TextUtils.writeLog("Payment Number is available in macanomy:" +PaymentNo);
 }
 else{ 
    ValidationUtils.verify(false,true,"Payment Number is Needed for Remittance Email");
  }
  
//  var paymentdate1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McDatePickerWidget;
//  Sys.HighlightObject(paymentdate1);
//      if((Paymentdate!="")&&(Paymentdate!=null)){
//       aqUtils.Delay(1000, Indicator.Text);
//          WorkspaceUtils.CalenderDateSelection(paymentdate1,Paymentdate)
//          ValidationUtils.verify(true,true,"Payment Date is selected in Maconomy"); 
//        }
//    else{ 
//      ValidationUtils.verify(false,true,"Payment Date is Needed  for Remittance Email");
//    } 
  
//  var paymentdate2 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McDatePickerWidget2;
//  Sys.HighlightObject(paymentdate2);
//      if((Paymentdate!="")&&(Paymentdate!=null)){
//       aqUtils.Delay(1000, Indicator.Text);
//          WorkspaceUtils.CalenderDateSelection(paymentdate2,Paymentdate)
//          ValidationUtils.verify(true,true,"Payment Date is selected in Maconomy"); 
//        }
//    else{ 
//      ValidationUtils.verify(false,true,"Payment Date is Needed for Remittance Email");
//    } 
//  
  var vendor1 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
    Sys.HighlightObject(vendor1);
    if((VendorNo!="")&&(VendorNo!=null)){
  vendor1.Click();
  WorkspaceUtils.VPWSearchByValue(vendor1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
    }
 else{ 
    ValidationUtils.verify(true,true,"Vendor Number is Needed for Remittance Email");
  }
  
  var vendor2 = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget2;
    Sys.HighlightObject(vendor2);
    if((VendorNo!="")&&(VendorNo!=null)){
  vendor2.Click();
  WorkspaceUtils.VPWSearchByValue(vendor2,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  TextUtils.writeLog("Vendor Number is available in macanomy: "+VendorNo);
    }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed for Remittance Email");
  }
  
  aqUtils.Delay(2000, Indicator.Text);
  var Do_Not_Show_Sent = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  if(!Do_Not_Show_Sent.getSelection()){ 
  Do_Not_Show_Sent.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  Do_Not_Show_Sent.Click();
  ReportUtils.logStep("INFO", "Do Not Show Sent is Checked");
  Log.Message("Do Not Show Sent is Checked")
  }
  
    aqUtils.Delay(2000, Indicator.Text); 
  var save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  waitForObj(save)  
  Sys.HighlightObject(save);
    save.Click();
    TextUtils.writeLog("Details are Saved"); 
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   
   var paymentOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
   paymentOrder.Click();
   aqUtils.Delay(2000, Indicator.Text);
   var flag= false;
   var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
   
     for(var v=0;v<table.getItemCount();v++){ 
       Log.Message(table.getItem(v).getText_2(1).OleValue.toString().trim()== EnvParams.Opco);
       Log.Message(table.getItem(v).getText_2(2).OleValue.toString().trim()==VendorNo);
       Log.Message((table.getItem(v).getText_2(7).OleValue.toString().trim()!="") || (table.getItem(v).getText_2(7).OleValue.toString().trim()!=null))
   if((table.getItem(v).getText_2(1).OleValue.toString().trim()== EnvParams.Opco) && 
  (table.getItem(v).getText_2(2).OleValue.toString().trim()==VendorNo) &&((table.getItem(v).getText_2(7).OleValue.toString().trim()!="") || (table.getItem(v).getText_2(7).OleValue.toString().trim()!=null))){  
  var CheckBox = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McPlainCheckboxView", "").SWTObject("Button", "");
    if(!CheckBox.getSelection()){ 
  CheckBox.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  CheckBox.Click();
  ReportUtils.logStep("INFO", "Payment Order is Selected");
  Log.Message("Payment Order is Selected")
  }
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
}
  
if(flag){
  var email = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
    Sys.HighlightObject(email);
    waitForObj(email);
  email.HoverMouse();
  ReportUtils.logStep_Screenshot("");
    email.Click();
  TextUtils.writeLog("Details are send to the Email"); 
   aqUtils.Delay(15000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  }
  
}



//Go To Job from Menu
function goToJobMenuItem(){

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
ImageRepository.ImageSet.AccountPayable2s.Click();
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
//Client_Managt.ClickItem("|Vendor Remittance E-mail");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Remittance E-mail").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Remittance E-mail").OleValue.toString().trim());
//Client_Managt.DblClickItem("|Vendor Remittance E-mail");
}
}
ReportUtils.logStep("INFO", "Moved to Banking Transactions from job Menu");
TextUtils.writeLog("Entering into Banking Transactions from Jobs Menu");
}






