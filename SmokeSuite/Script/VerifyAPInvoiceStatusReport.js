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
function verifyAPInvoiceStatus_Report(){
TextUtils.writeLog("Verification Of AP Invoice Status Report"); 

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
WorkspaceUtils.closeAllWorkspaces();
gotoMenu(); 
gotoAPScreen();
gotoAPInvoiceStatusLink();
verifyAPInvoiceStatusScreen();
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
scroll.MouseWheel(-500);

if(ImageRepository.ImageSet.Reporting.Exists()){
ImageRepository.ImageSet.Reporting.Click();// GL
}
else{
ImageRepository.ImageSet.Reporting_1.Click();
}

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|Reports");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Reports");
}

} 
ReportUtils.logStep("INFO", "Moved to Reports from Reporting Menu");
Log.Message("Entering into Reports from Reporting Menu");
}

function gotoAPScreen(){ 

var apScreen = Aliases.Maconomy.Reports_GL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(apScreen);
Sys.HighlightObject(apScreen);
apScreen.Click();
ReportUtils.logStep("INFO", "Clicked AP Menu Section");
Log.Message("Clicked AP Menu Section")
}

function gotoAPInvoiceStatusLink()
{
var apInvoiceLink = Aliases.Maconomy.Reports_GL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite4.McLinkLabelWidget.McTextWidget;
waitForObj(apInvoiceLink);
Sys.HighlightObject(apInvoiceLink);
ReportUtils.logStep_Screenshot();
apInvoiceLink.Click();
ReportUtils.logStep("INFO", "Clicked on AP Invoice Status Link");
Log.Message("Clicked on AP Invoice Status Link");
}


function verifyAPInvoiceStatusScreen()
{
  if(ImageRepository.Browser_Reporting.Browser_DataProtection_Dialog.Exists())
    ImageRepository.Browser_Reporting.Browser_DataProtection_OK_Button.Click();

  if(ImageRepository.Browser_Reporting.Browser_GLTransaction_Prompt.Exists())
    ImageRepository.Browser_Reporting.Browser_GLTransaction_Prompt_Cancel.Click();
  
  if(ImageRepository.Browser_Reporting.APInvoiceStatus_Logo.Exists())
  {
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "AP Invoice Status Screen displayed sucessfully");
     Log.Message("AP Invoice Status Screen displayed sucessfully");
     } 
  else
     ReportUtils.logStep("Fail", "AP Invoice Status Screen not displayed");            
}