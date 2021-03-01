//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

/** 
 * This script implements verification of Balance sheet screen in a Web Browser
 * @author  : Sai Kiran Vemula
 * @version : 1.0
 * Created Date :07/31/2020
 */
var Project_manager = "";
var Language = "";

 /**
  *  This function invokes maconomy and calls subfunctionality methods
  */
function verifyBalanceSheet_Report(){
TextUtils.writeLog("Verification Of Balance Sheet Report"); 

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

try{
//Entering Report Management
WorkspaceUtils.closeAllWorkspaces();
gotoMenu(); 
gotoBalanceSheetLink();
verifyBalanceSheetScreen();
}
  catch(err){
    Log.Message(err);
  }
}

/**
  *  This function navigates to Statutory Reports
  */
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
report_mngmnt.ClickItem("|Statutory Reports");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Statutory Reports");
}

} 
ReportUtils.logStep("INFO", "Moved to Statutory Reports from Reporting Menu");
Log.Message("Entering into Statutory Reports from Reporting Menu");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{}
}


/**
  *  This function clicks Balance Sheet link in Statutory Reports
  */
function gotoBalanceSheetLink()
{
var balanceSheetLink = Aliases.Maconomy.Reports_GL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McLinkLabelWidget.McTextWidget;
waitForObj(balanceSheetLink);
Sys.HighlightObject(balanceSheetLink);
ReportUtils.logStep_Screenshot();
balanceSheetLink.HoverMouse();
balanceSheetLink.Click();
ReportUtils.logStep("INFO", "Clicked on Balance Sheet Link");
Log.Message("Clicked on Balance Sheet Link");
}


/**
  *  This function is to verify Screen on Web Browser
  */

function verifyBalanceSheetScreen()
{
  if(ImageRepository.Browser_Reporting.Browser_DataProtection_Dialog.Exists())
    ImageRepository.Browser_Reporting.Browser_DataProtection_OK_Button.Click();

    aqUtils.Delay(3000,"Waiting for Prompt window");
  if(ImageRepository.Browser_Reporting.Browser_GLTransaction_Prompt.Exists())
    ImageRepository.Browser_Reporting.Browser_GLTransaction_Prompt_Cancel.Click();
    
 var pageName = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.frameIframeleftpanew.cell.panelDivdocname.textContent;

  
  if(pageName.trim() == "Balance Sheet" || ImageRepository.Browser_Reporting.BalanceSheet_Logo.Exists())
  {
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Balance Sheet Screen displayed sucessfully");
     Log.Message("Balance Sheet Screen displayed successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Balance Sheet Screen not displayed"); 
     
  Sys.Browser("chrome").BrowserWindow(0).Keys("^w");    
}