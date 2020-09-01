//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

//Indicator.Show();
var Project_manager = "";
var Language = "";
//Strating Of TestCase
function verifyGeneralLedger(){
TextUtils.writeLog("Verification Of General Ledger Worksapce"); 

//Setting Language in WorkspaceUtils
Language = "";
Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

//Checking Login for Client Creation
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
}

//Initializing Variables

try{
WorkspaceUtils.closeAllWorkspaces();
gotoMenu(); 
verifyGLTransactions();
WorkspaceUtils.closeAllWorkspaces();
gotoGLSetupWorkspace();
verifyGLSetupWorkspace();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.GeneralLedger.Exists()){
ImageRepository.ImageSet.GeneralLedger.Click();
}
else{
     ReportUtils.logStep("Fail", "General Ledger Workspace not displayed");
     Log.Message("General Ledger not displayed");
}

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
Sys.HighlightObject(Workspc);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|GL Transactions");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|GL Transactions");
}
}
aqUtils.Delay(3000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to General Ledger Transactions Workspace from General Ledger Menu");
Log.Message("Entering General Ledger Transactions Workspace from General Ledger Menu");
}
}

function verifyGLTransactions(){ 
aqUtils.Delay(5000, "Waiting to Load");
var glTransactionsTab= Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(glTransactionsTab);
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var generalJournalTab = Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(generalJournalTab);
if(generalJournalTab.isVisible())
{
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "GL Transactions are loaded successfully");
     Log.Message("GL Transactions are loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "GL Transactions are not loaded");
}

function gotoGLSetupWorkspace(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.GeneralLedger.Exists()){
ImageRepository.ImageSet.GeneralLedger.Click();
}
else{
     ReportUtils.logStep("Fail", "General Ledger Section not displayed");
     Log.Message("General Ledger Section not displayed");
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

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|GL Setup");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|GL Setup");
} 
} 
aqUtils.Delay(2000, "Waiting to Load");

if(ImageRepository.ImageSet.ExcelImage.Exists())
  {
  ReportUtils.logStep("INFO", "Moved to GL Setup from General Ledger Menu");
  Log.Message("Entering into GL Setup from General Ledger Menu");
  }
}

function verifyGLSetupWorkspace(){ 
aqUtils.Delay(5000, "Waiting to Load");
var glSetupTab=Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl

Sys.HighlightObject(glSetupTab);
glSetupTab.Click();

var listOfComp = Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(listOfComp);
listOfComp.Click();

var companyNoSearch = Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget
companyNoSearch.Click();
companyNoSearch.setText(EnvParams.Opco);
if(ImageRepository.ImageSet.RefreshIcon.Exists()){}
companyNoSearch.Keys("[Down]");
Sys.Keys("[Enter]");
if(ImageRepository.ImageSet.RefreshIcon.Exists()){}

var informationTab = Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
waitForObj(informationTab);
Sys.HighlightObject(informationTab);
informationTab.Click();
aqUtils.Delay(2000, "Waiting to Load");
if(ImageRepository.ImageSet.RefreshIcon.Exists()){}

var companyNumber = Aliases.Maconomy.GeneralLedger.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget
if(EnvParams.Opco == companyNumber.getText().trim())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(companyNumber);
     ReportUtils.logStep("Pass", "GL Setup Information loaded successfully");
     Log.Message("GL Setup Information loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "GL Setup Information is not loaded");
}
