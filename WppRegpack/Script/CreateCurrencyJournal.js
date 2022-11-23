//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateCurrencyJournal";

ager="";
var level =0;
var STIME = "";
var ReversionDate = "";
var Language = "";
//Main Function

function CurrencyJournal(){ 
TextUtils.writeLog("Create Currency Journal Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Senior Accountant","Username");
if((Project_manager=="")||(Project_manager==null))
ValidationUtils.verify(false,true,"Login Credentials required for anyone of SSC - Junior Accountant or SSC - Senior Accountant");

Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
 
}
getDetails();
gotoMenu();
gotToCurrencyRevaluation();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.GendralLedger.Exists()){
ImageRepository.ImageSet.GendralLedger.Click();// GL
}
else if(ImageRepository.ImageSet.GendralLedger1.Exists()){
ImageRepository.ImageSet.GendralLedger1.Click();
}
else{
ImageRepository.ImageSet.GendralLedger2.Click();
}


var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Transactions from General Ledger Menu");
TextUtils.writeLog("Entering into GL Transactions from General Ledger Menu");
}


function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Entrydate = ExcelUtils.getRowDatas("Reversion Date",EnvParams.Opco)
if (Entrydate == "AUTOFILL")
  Entrydate = getSpecificDate(1)
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"Reversion Date is Needed to Create Currency Journal");
}
Log.Message(Entrydate)


layoutTypes = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(layoutTypes)
if((layoutTypes==null)||(layoutTypes=="")){ 
ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Selection");
}

}

function gotToCurrencyRevaluation(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var CurrencyReval = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(CurrencyReval);
CurrencyReval.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var CompanyNo = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
WorkspaceUtils.waitForObj(CompanyNo);
CompanyNo.setText(EnvParams.Opco)
var CompanyNo = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget2
WorkspaceUtils.waitForObj(CompanyNo);
CompanyNo.setText(EnvParams.Opco)
var GLEntries = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McPlainCheckboxView.Button
WorkspaceUtils.waitForObj(GLEntries);
GLEntries.Click();

var RDate = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite2.McDatePickerWidget
WorkspaceUtils.waitForObj(RDate);
//WorkspaceUtils.CalenderDateSelection(RDate,Entrydate)
RDate.setText(Entrydate);
aqUtils.Delay(2000, "Saving the Changes");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var layout = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite2.McPopupPickerWidget;
layout.Keys(layoutTypes);
  Delay(5000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var StatementDate = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;
if((StatementDate.getText()=="")||(StatementDate.getText()==null)){ 
  StatementDate.Click();
  StatementDate.setText( getSpecificDate(0) );

}
aqUtils.Delay(2000, "Saving the Changes");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var Save = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
WorkspaceUtils.waitForObj(Save);
Save.Click();
aqUtils.Delay(2000, "Saving the Changes");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){   
}
var print = Aliases.Maconomy.CurrencyJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2
WorkspaceUtils.waitForObj(print);
print.Click();
while(!ImageRepository.ImageSet.Tab_Icon.Exists()){   }
   TextUtils.writeLog("Post and Email is Clicked");
    //aqUtils.Delay(5000, Indicator.Text);

    WorkspaceUtils.savePDF_To_LocalDirectory("PDF Currency Report","P_CurrencyReport");

}