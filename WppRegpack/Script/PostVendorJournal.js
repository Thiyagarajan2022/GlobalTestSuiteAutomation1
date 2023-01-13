﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT EventHandler


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

function postVendorJournal(){ 
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

  gotoMenu();
  Delay(5000);
searchForJournal();
postJournal();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}



function postVendorJournal_Dependency(){ 
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

  gotoMenu();
  Delay(5000);
searchForJournal();
postJournal();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());
 TextUtils.writeLog(" Navigate AP Transactions "); 
     ReportUtils.logStep_Screenshot("");
}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable Menu");

 
 
}

function postJournal(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var postVendotJournal = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(postVendotJournal);
postVendotJournal.click();

aqUtils.Delay(1000, Indicator.Text);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Post Vendor Journal",EnvParams.Opco,"Data Management",JournalNo);

aqUtils.Delay(1000, Indicator.Text);
WorkspaceUtils.savePDF_To_LocalDirectory("PDF Print Posting Journal","Print Posting Journal");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Vendor Invoice Journal").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions - Vendor Invoice Journal").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
var OKay = Aliases.Maconomy.GLJornalAwaitingApproval.Okay.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OKay.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}
  

function searchForJournal(){ 
aqUtils.Delay(2000, "finding Company field");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Sys.Desktop.Keys("[Up]");
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







