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
function verifyPersonalInfo(){
TextUtils.writeLog("Verification Of Personal Information Workspace"); 

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
gotoCardBankDetailsScreen();
verifyEmpPaymentInfoScreen();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();


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

var scroll = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "")
scroll.Click();
scroll.MouseWheel(500);
if(ImageRepository.ImageSet.PersonalInformation.Exists()){
ImageRepository.ImageSet.PersonalInformation.Click();
}
else{
     ReportUtils.logStep("Fail", "Personal Info Section not displayed");
     Log.Message("Personal Info Section not displayed");
}

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|Card and Bank Details");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Card and Bank Details");
}

} 

if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Card and Bank Details from Personal Informaion Menu");
Log.Message("Entering into Card and Bank Details from Personal Informaion Menu");
}
}

function gotoCardBankDetailsScreen(){ 
aqUtils.Delay(5000, "Waiting to Load");
var cardDetailsScreen = Aliases.Maconomy.PersonalInformation.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(cardDetailsScreen);
cardDetailsScreen.Click();
ReportUtils.logStep("INFO", "Clicked Card and Bank Details Tab");
Log.Message("Clicked Card and Bank Details Tab")
}

function verifyEmpPaymentInfoScreen()
{
aqUtils.Delay(5000, "Waiting to Load");
var empInfoSection = Aliases.Maconomy.PersonalInformation.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(empInfoSection);
ReportUtils.logStep_Screenshot();
empInfoSection.Click();
ReportUtils.logStep("INFO", "Clicked on EmpInfo Section");
Log.Message("Clicked on Empo Info");

var cardAndBankDetailsSection = Aliases.Maconomy.PersonalInformation.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget;

if(cardAndBankDetailsSection.isVisible())
{
     Sys.HighlightObject(empInfoSection);
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Bank details section displayed sucessfully");
     Log.Message("Bank details section displayed sucessfully");
     } 
  else
     ReportUtils.logStep("Fail", "Bank details section not displayed");             

}

