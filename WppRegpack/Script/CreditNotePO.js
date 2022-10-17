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
var Ipreparation_ID,IpreparationUnit,JobNoTo = "";
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


//ExcelUtils.setExcelName(workBook, "Data Management", true);
//POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
//if((POnumber==null)||(POnumber=="")){ 
//TextUtils.writeLog("Creating Negative Purchase Order for Credit Note");
//CreatePO = true;
//Runner.CallMethod("CreatePurchaseOrder.CreatePurchaseOrder");
//}
//ExcelUtils.setExcelName(workBook, "Data Management", true);
//POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
//Log.Message(POnumber)
//CreatePO = false;
//ExcelUtils.setExcelName(workBook, "Data Management", true);
//Approved_POnumber = ReadExcelSheet("Approved Credit PO Number",EnvParams.Opco,"Data Management");
//if((Approved_POnumber==null)||(Approved_POnumber=="")){ 
//TextUtils.writeLog("Approving Negative Purchase Order for Credit Note");
//CreatePO = true;
//Runner.CallMethod("ApprovePurchaseOrder.ApprovePurchaseOrder");
//}
//
//CreatePO = false;


   Get_Details();
   Login_Maconomy();
   goToAPMenuItem(); 
   invoiceAllocation();

WorkspaceUtils.closeAllWorkspaces();
}






function Login_Maconomy(){
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

function Get_Details(){
sheetName = "CreditNoteWithPO";
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("Approved Credit PO Number",EnvParams.Opco,"Data Management");
if((POnumber=="")||(POnumber==null)){

Create_New_Job_2();

}
Log.Message("Get Details POnumber :"+POnumber)
ExcelUtils.setExcelName(workBook, sheetName, true);

company = EnvParams.Opco;
//company = ExcelUtils.getColumnDatas("Opco",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Vendor Invoice");
}
InvoiceNo = ExcelUtils.getColumnDatas("Invoice No",EnvParams.Opco)
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Invoice No. is Needed to Create a Vendor Invoice");
}
Log.Message(InvoiceNo)
EDate = ExcelUtils.getColumnDatas("Entry Date",EnvParams.Opco)
if(EDate == "AUTOFILL")
        EDate = getSpecificDate(0);   
if((EDate==null)||(EDate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Vendor Invoice");
}
Log.Message(EDate)
IDate = ExcelUtils.getColumnDatas("Invoice Date",EnvParams.Opco)
if(IDate == "AUTOFILL")
        IDate = getSpecificDate(0);
if((IDate==null)||(IDate=="")){ 
ValidationUtils.verify(false,true,"Invoice Date is Needed to Create a Vendor Invoice");
}
Log.Message(IDate)
description = ExcelUtils.getColumnDatas("Description1",EnvParams.Opco)
if((description==null)||(description=="")){ 
ValidationUtils.verify(false,true,"Description1 is Needed to Create a Vendor Invoice");
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
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl
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

Log.Message("Invoice POnumber :"+POnumber)
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
    CreatePO = true;
    Runner.CallMethod("CreatePO.CreatePurchaseOrder",POSheet,JobSO,PO_SO);
    Log.PopLogFolder();
    CreatePO = false;
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    }
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    PoNumber2 = ExcelUtils.getRowDatas("PO Number_"+PO_SO,EnvParams.Opco)
    POnumber = PoNumber2;
    Log.Message("POnumber :"+POnumber)
    Log.Message("POnumber :"+PoNumber2)
    
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
    CreatePO = true;
    Runner.CallMethod("ApprovePO.ApprovePurchaseOrder",APSheet,serialOder);
    Log.PopLogFolder();
    CreatePO = false;
    ReportUtils.Dreport.endTest(ReportUtils.Dtest);
    ReportUtils.Dreport.flush();
    Runner.CallMethod("JIRA.JIRAUpdate");
    ReportUtils.DStat = false;
    Approved_POnumber = PoNumber2;
    Log.Message("Approved_POnumber :"+Approved_POnumber)
   }

TestRunner.testCaseId = Ipreparation_ID;
TestRunner.unitName = IpreparationUnit;

Log.Message("Job To :"+JobNoTo)



TestRunner.JiraStat = true;
TestRunner.JiraUpdate = true;

}



