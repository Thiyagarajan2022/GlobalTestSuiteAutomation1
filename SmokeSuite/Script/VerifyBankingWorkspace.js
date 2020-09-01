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
function verifyBankingTransactions(){
TextUtils.writeLog("Verification Of Accounts Receivable Workspace"); 

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
verifyBankTransactions();
WorkspaceUtils.closeAllWorkspaces();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.BankingWorkspace.Exists()){
ImageRepository.ImageSet.BankingWorkspace.Click();
}
else if(ImageRepository.ImageSet.BankingWorkspace_2.Exists()){
ImageRepository.ImageSet.BankingWorkspace_2.Click();
}
else{
     ReportUtils.logStep("Fail", "Banking Workspace not displayed");
     Log.Message("Banking Workspace not displayed");
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
report_mngmnt.ClickItem("|Bank Transactions");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Bank Transactions");
}
}
aqUtils.Delay(2000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Bank Transactions workspace from Banking Menu");
Log.Message("Entering Bank Transactions workspace from Banking Menu");
}
}

function verifyBankTransactions(){ 
aqUtils.Delay(3000, "Waiting to Load");
var bankTransactionsTab= Aliases.Maconomy.BankingWorkspace.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
Sys.HighlightObject(bankTransactionsTab);

var bankPaymentsTab = Aliases.Maconomy.BankingWorkspace.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(bankPaymentsTab);

var selectionTab = Aliases.Maconomy.BankingWorkspace.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(selectionTab);

var createTab = Aliases.Maconomy.BankingWorkspace.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(createTab);

var selectionCriteria = Aliases.Maconomy.BankingWorkspace.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget;

if(selectionCriteria.isVisible())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(selectionCriteria);
     ReportUtils.logStep("Pass", "Create Payment Selection Workspace loaded successfully");
     Log.Message("Create Payment Selection Workspace loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Create Payment Selection Workspace not loaded");
}


