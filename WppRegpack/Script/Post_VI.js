//USEUNIT EnvParams
//USEUNIT EventHandler
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "";
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

function postVendorJournal(sheet, VIno){ 
Indicator.PushText("waiting for window to open");
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
sheetName = sheet;
VInum = VIno;
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

//ExcelUtils.setExcelName(workBook, sheetName, true);
//var fName = ExcelUtils.getColumnDatas("JIRA Opco Name",EnvParams.Opco)
//if((fName=="")||(fName==null))
//ValidationUtils.verify(false,true,"JIRA Opco Name is Needed to update status of Create a Purchase Order");
//else{ 
//EventHandler.folderName = fName;
//}
//
//var TestID = ExcelUtils.getColumnDatas("JIRA TestCase ID",EnvParams.Opco)
//if((TestID=="")||(TestID==null))
//ValidationUtils.verify(false,true,"JIRA TestCase ID is Needed to update status of Create a Purchase Order");
//else{ 
//EventHandler.testCaseId = TestID; 
//}

//VendorID = ExcelUtils.getColumnDatas("Vendor Number",EnvParams.Opco)
//  if((VendorID=="")||(VendorID==null)){
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  VendorID = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
//  }
//if((VendorID==null)||(VendorID=="")){ 
//ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Purchase Order");
//}
//

//JournalNo = ExcelUtils.getColumnDatas("journal No",EnvParams.Opco)
//  if((JournalNo=="")||(JournalNo==null)){
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  JournalNo = ReadExcelSheet("journal No",EnvParams.Opco,"Data Management");
//  }
//if((JournalNo==null)||(JournalNo=="")){ 
//ValidationUtils.verify(false,true,"Job Number is Needed to Create a Purchase Order");
//}


  ExcelUtils.setExcelName(workBook, "Data Management", true);
  JournalNo = ReadExcelSheet("Invoice Journal NO_"+VInum,EnvParams.Opco,"Data Management");
  if((JournalNo=="")||(JournalNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  JournalNo = ExcelUtils.getRowDatas("Journal No",EnvParams.Opco)
  }
  if((JournalNo=="")||(JournalNo==null))
  ValidationUtils.verify(false,true,"journal No is required to create USER");

//JournalNo = ExcelUtils.getRowDatas("journal No",EnvParams.Opco)
//if((JournalNo==null)||(JournalNo=="")){ 
//ValidationUtils.verify(false,true,"journal No is required to create USER");
//}

companyNo = EnvParams.Opco
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

  gotoMenu();
  Delay(5000);
searchForJournal();
postJournal();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();

 TextUtils.writeLog(" Post Vendor Journal Started "); 
     ReportUtils.logStep_Screenshot("");
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
aqUtils.Delay(3000, Indicator.Text);
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
Client_Managt.ClickItem("|AP Transactions");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|AP Transactions");
 TextUtils.writeLog(" Navigate AP Transactions "); 
     ReportUtils.logStep_Screenshot("");
}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable Menu");

 
 
}

function postJournal(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var postVendotJournal = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(postVendotJournal);
postVendotJournal.click();

//Delay(20000);
aqUtils.Delay(5000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Post Vendor Journal_"+VInum,EnvParams.Opco,"Data Management",JournalNo);

var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Posting Journal")!=-1){
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
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
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

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Vendor Invoice Journal").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Vendor Invoice Journal").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 

//Sys.Process("Maconomy").SWTObject("Shell", "AP Transactions - Vendor Invoice Journal").SWTObject("Label", "Maconomy cannot show more lines. The table cannot be modified")
//Sys.Process("Maconomy").SWTObject("Shell", "AP Transactions - Vendor Invoice Journal").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
var OKay = Aliases.Maconomy.GLJornalAwaitingApproval.Okay.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OKay.Click();
}

//var SaveTitle = "";
//var sFolder = "";
//var pdf =    Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "print posting journal.pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5)
//   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "print posting journal.pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("print posting journal")!=-1){
//
////Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "print job quote"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
////    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "print job quote"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("print job quote")!=-1){
// 
// TextUtils.writeLog("Post Vendor Journal"); 
//     ReportUtils.logStep_Screenshot("");
//
//Sys.HighlightObject(pdf);
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//    
//    
//ValidationUtils.verify(true,true,"Print is Generated");
//Log.Message("Print is Generated")
//ReportUtils.logStep("INFO","Print is Generated");

    
//  aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x41);
//    
//        
//    if(ImageRepository.PDF.ChooseFolder.Exists())
//    ImageRepository.PDF.ChooseFolder.Click();
//    
//    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    aqUtils.Delay(2000, Indicator.Text);
//    SaveTitle = save.wText;
//    
//sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
//if (! aqFileSystem.Exists(sFolder)){
//if (aqFileSystem.CreateFolder(sFolder) == 0){ 
//    
//}
//else{
//Log.Error("Could not create the folder " + sFolder);
//}
//}
//save.Keys(sFolder+SaveTitle+".pdf");
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
//aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
//Sys.HighlightObject(pdf);
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//    }
//ValidationUtils.verify(true,true,"Print Draft Quote is Clicked and PDF is Saved");
//Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
//ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

//}
}
  

function searchForJournal(){ 
aqUtils.Delay(2000, "finding Company field");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy Loading Data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var companyNoField = Aliases.ObjectGroup.CompanyNoVendorJournal;

companyNoField.setText(companyNo)
companyNoField.Keys("[Tab]"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var journalNo = Aliases.ObjectGroup.JournalNoField;
journalNo.Keys(JournalNo);

aqUtils.Delay(5000, "Searching Journal in Table");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 TextUtils.writeLog("Post Vendor Journal"); 
 ReportUtils.logStep_Screenshot("");
 
  
  var closefilter =  Aliases.ObjectGroup.CloseFilterVendorJournal
closefilter.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


}







