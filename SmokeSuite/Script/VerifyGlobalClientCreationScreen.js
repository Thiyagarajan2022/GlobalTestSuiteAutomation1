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
function verifyGlobalClientCreation(){
TextUtils.writeLog("Verification Of New Global Client Creation"); 

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
verifyClientCreation();
WorkspaceUtils.closeAllWorkspaces();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountsReceivable.Exists()){
ImageRepository.ImageSet.AccountsReceivable.Click();
}
else{
     ReportUtils.logStep("Fail", "AccountsReceivable Workspace not displayed");
     Log.Message("AccountsReceivable Workspace not displayed");
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
report_mngmnt.ClickItem("|Client Management");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Client Management");
}
}
aqUtils.Delay(2000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Client Management workspace from Accounts Receivable Menu");
Log.Message("Entering Client Management workspace from Accounts Receivable Menu");
}
}

function verifyClientCreation(){ 

var clientMangementTab= Aliases.Maconomy.AR_ClientManagement.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(clientMangementTab);

Sys.HighlightObject(clientMangementTab);

var globalClientsTab = Aliases.Maconomy.AR_ClientManagement.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(globalClientsTab);
globalClientsTab.Click();

ImageRepository.ImageSet.RefreshIcon.Click();
aqUtils.Delay(2000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var newGlobalClientButton = Aliases.Maconomy.AR_ClientManagement.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;

Sys.HighlightObject(newGlobalClientButton);
newGlobalClientButton.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}


var globalClientMasterDataLabel = Aliases.Maconomy.Shell.Composite.Label;
//Aliases.Maconomy.AR_ClientMangmnt_GlobalClient.Composite.Label;

if(globalClientMasterDataLabel.isVisible())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(globalClientMasterDataLabel);
     ReportUtils.logStep("Pass", "New Global Client Screen loaded successfully");
     Log.Message("New Global Client Screen loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "New Global Client Screen not loaded");


var globalClientMasterCancelBttn = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Button;
//Aliases.Maconomy.AR_ClientMangmnt_GlobalClient.Composite.Composite.Composite.Composite.Button;
globalClientMasterCancelBttn.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}

}


