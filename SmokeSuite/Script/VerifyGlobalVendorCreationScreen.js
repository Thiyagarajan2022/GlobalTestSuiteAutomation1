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
function verifyGlobalVendorCreation(){
TextUtils.writeLog("Verification Of New Global Vendor Creation"); 

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
//Entering Report Management
gotoMenu(); 
verifyVendorCreation();
WorkspaceUtils.closeAllWorkspaces();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountsPayable.Exists())
ImageRepository.ImageSet.AccountsPayable.Click();
else if (ImageRepository.ImageSet.AccountsPayable_2.Exists())
ImageRepository.ImageSet.AccountsPayable_2.Click();
else{
     ReportUtils.logStep("Fail", "Accounts Payable Workspace not displayed");
     Log.Message("Accounts Payable Workspace not displayed");
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
report_mngmnt.ClickItem("|Vendor Management");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Vendor Management");
}
}
aqUtils.Delay(2000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Vendor Management workspace from Accounts Payable Menu");
Log.Message("Entering Vendor Management workspace from Accounts Payable Menu");
}
}

function verifyVendorCreation(){ 

var vendorManagementTab= Aliases.Maconomy.AP_VendorMangmt_GlobalVendor.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(vendorManagementTab);
Sys.HighlightObject(vendorManagementTab);

var globalVendorsTab = Aliases.Maconomy.AP_VendorMangmt_GlobalVendor.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(globalVendorsTab);
Sys.HighlightObject(globalVendorsTab);
globalVendorsTab.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}
ImageRepository.ImageSet.RefreshIcon.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var newGlobalVendorButton = Aliases.Maconomy.AP_VendorMangmt_GlobalVendor.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;

Sys.HighlightObject(newGlobalVendorButton);
newGlobalVendorButton.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}


// Address has to be captured Again

var globalVendorMasterDataLabel = Aliases.Maconomy.AR_ClientMangmnt_GlobalClient.Composite.Label;

if(globalVendorMasterDataLabel.isVisible())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(globalVendorMasterDataLabel);
     ReportUtils.logStep("Pass", "New Global Vendor Screen loaded successfully");
     Log.Message("New Global Vendor Screen loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "New Global Vendor Screen not loaded");


var globalVendorMasterCancelBttn = Aliases.Maconomy.AR_ClientMangmnt_GlobalClient.Composite.Composite.Composite.Composite.Button;
globalVendorMasterCancelBttn.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}

}


