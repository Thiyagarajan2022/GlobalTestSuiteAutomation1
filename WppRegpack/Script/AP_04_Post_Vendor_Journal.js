﻿//USEUNIT EnvParams
//USEUNIT EventHandler
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT ObjectUtils
//USEUNIT ActionUtils

/** 
 * This script create Post Vendor Invoice
 * @author  : Muthu Kumar M
 * @version : 3.0
  * Modified Date(MM/DD/YYYY) : 02/21/2022
 */


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "PostVendorJournal";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
//var VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice = "";
var companyNo,JournalNo;
var Language= "";
var VInum = "";
var Maconomy_ParentAddress,Maconomy_Index = "";

function postVendorJournal(){ 
Indicator.PushText("waiting for window to open");

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Approve Vendor Invoice
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Senior AP","Username")

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);




excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "PostVendorJournal";
VInum = "";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
//VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice = "";
companyNo,JournalNo ="";
  
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  JournalNo = ReadExcelSheet("Reverse CreditNote Invoice Journal NO",EnvParams.Opco,"Data Management");
  if((JournalNo=="")||(JournalNo==null)){
    
  JournalNo = ReadExcelSheet("CreditNote Invoice Journal NO",EnvParams.Opco,"Data Management");
  if((JournalNo=="")||(JournalNo==null)){
    
  JournalNo = ReadExcelSheet("Reverse Invoice Journal NO",EnvParams.Opco,"Data Management");
  if((JournalNo=="")||(JournalNo==null)){
    
  JournalNo = ReadExcelSheet("Invoice Journal NO",EnvParams.Opco,"Data Management");
  if((JournalNo=="")||(JournalNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  JournalNo = ExcelUtils.getRowDatas("Journal No",EnvParams.Opco)
  }
  if((JournalNo=="")||(JournalNo==null))
  ValidationUtils.verify(false,true,"journal No is required to create USER");
  
  else{ 
  ValidationUtils.verify(true,true,"Posting Invoice Journal NO :"+JournalNo)
}

}
else{ 
  ValidationUtils.verify(true,true,"Posting Reverse Invoice Journal NO :"+JournalNo)
}

}
else{ 
  ValidationUtils.verify(true,true,"Posting CreditNote Invoice Journal NO :"+JournalNo)
}

}
else{ 
  ValidationUtils.verify(true,true,"Posting Reverse CreditNote Invoice Journal NO :"+JournalNo)
}
companyNo = EnvParams.Opco
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

  goTo_Account_Payable();
  Delay(5000);
  searchForJournal();
  postJournal();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}



function postVendorJournal_Dependency(){ 
Indicator.PushText("waiting for window to open");
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();

//Checking Login to execute Approve Vendor Invoice
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Senior AP","Username")

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);


excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "PostVendorJournal";
VInum = "";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
//VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice = "";
companyNo,JournalNo ="";
  
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "PO Creation started::"+STIME);

  ExcelUtils.setExcelName(workBook, "Data Management", true);
    
  JournalNo = ReadExcelSheet("Second Invoice Journal NO",EnvParams.Opco,"Data Management");
  if((JournalNo=="")||(JournalNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  JournalNo = ExcelUtils.getRowDatas("Journal No",EnvParams.Opco)
  }
  if((JournalNo=="")||(JournalNo==null))
  ValidationUtils.verify(false,true,"journal No is required to create USER");
  
  else{ 
  ValidationUtils.verify(true,true,"Posting Invoice Journal NO :"+JournalNo)
}



companyNo = EnvParams.Opco
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

  goTo_Account_Payable();
  Delay(5000);
  searchForJournal();
  postJournal();

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}

function goTo_Account_Payable(){ 
  
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_AccountPayable_from_workspace(); //Select Account Payable Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());

ReportUtils.logStep("INFO", "Moved to AP Transactions from Accounts Payable Menu");
TextUtils.writeLog("Entering into AP Transactions from Accounts Payable Menu");
}





function postJournal(){ 

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
var postVendotJournal = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Post");
//var postVendotJournal = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(postVendotJournal);
postVendotJournal.click();

//Delay(20000);
aqUtils.Delay(5000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Post Vendor Journal",EnvParams.Opco,"Data Management",JournalNo);

aqUtils.Delay(5000, Indicator.Text);


var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*", 1).WndCaption.indexOf("Print Posting Journal")!=-1){
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
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
//aqUtils.Delay(2000, Indicator.Text);

var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
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

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Vendor Invoice Journal").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Vendor Invoice Journal").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
var OKay = Aliases.Maconomy.GLJornalAwaitingApproval.Okay.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OKay.Click();
}

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


}
  

function searchForJournal(){ 

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
companyNoField = table.SWTObject("McValuePickerWidget", "")


companyNoField.setText(companyNo)
companyNoField.Keys("[Tab]"); 

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var journalNo = table.SWTObject("McTextWidget", "", 2)
Aliases.ObjectGroup.JournalNoField;
journalNo.Keys(JournalNo);

aqUtils.Delay(5000, "Searching Journal in Table");
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

 TextUtils.writeLog("Post Vendor Journal"); 
 ReportUtils.logStep_Screenshot("");
 
  
  var closefilter =  ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
closefilter.Click();


ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


}







