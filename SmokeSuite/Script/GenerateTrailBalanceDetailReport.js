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
function generateTrailBalanceDetail(){
TextUtils.writeLog("Generation Of Trail Balance Detail Report"); 

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
gotoGLScreen();
gotoTrailBalanceDetail();
generateTrailBalanceDetailReport();
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
Log.Message("Entering into GL Reports from Reporting Menu");
}

function gotoGLScreen(){ 
var glScreen = Aliases.Maconomy.Reports_GL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.GLTab;
waitForObj(glScreen);
Sys.HighlightObject(glScreen);
glScreen.Click();
ReportUtils.logStep("INFO", "Clicked GL Menu Section");
Log.Message("Clicked GL Menu Section")
}

function gotoTrailBalanceDetail()
{
var glTrailBalanceDetailLink = Aliases.Maconomy.Reports_GL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McLinkLabelWidget.McTextWidget;
waitForObj(glTrailBalanceDetailLink);
Sys.HighlightObject(glTrailBalanceDetailLink);
ReportUtils.logStep_Screenshot();
glTrailBalanceDetailLink.Click();
ReportUtils.logStep("INFO", "Clicked on Trail Balance Detail Link");
Log.Message("Clicked on Trail Balance Detail Link");
}


function generateTrailBalanceDetailReport()
{
var excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "TrailBalanceDetailReport";
ExcelUtils.setExcelName(workBook, sheetName, true);
var company =  ExcelUtils.getRowDatas("Company",EnvParams.Opco);


aqUtils.Delay(5000, "Navigating to Browser");
   
if(ImageRepository.Browser_Reporting.Browser_DataProtection_Dialog.Exists())
    ImageRepository.Browser_Reporting.Browser_DataProtection_OK_Button.Click();


  if(ImageRepository.Browser_Reporting.Browser_GLTransaction_Prompt.Exists())
    {
      
    var companySearch = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.textboxLovwidgetpromptlovzoneSea;
    waitForObj(companySearch);
    companySearch.SetText(company);
    
    var searchIcon = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelIconimgIconmenuIconLovwidge;
    waitForObj(searchIcon);
    searchIcon.Click();
    
    var searchResult = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelMclm0;

    Log.Message("Company is Listed in Search Result")
    searchResult.Click();
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellBtncimgLrzArrowaddPromptlovz;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();
    
    var periodFrom = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cell5;
    periodFrom.Click();
    var fromPeriod =  ExcelUtils.getRowDatas("PeriodFrom",EnvParams.Opco);
    var startDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.textboxLovwidgetpromptlovzoneTex;
    waitForObj(startDateField);
    startDateField.setText(fromPeriod);
    
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellBtncimgLrzArrowaddPromptlovz;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();
    
    
    var periodTo = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cell6;
    periodTo.Click();
    var toPeriod = ExcelUtils.getRowDatas("PeriodTo",EnvParams.Opco);
    var endDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.textboxLovwidgetpromptlovzoneTex;
    waitForObj(endDateField);
    endDateField.setText(toPeriod);
    
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellBtncimgLrzArrowaddPromptlovz;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();

    ReportUtils.logStep("INFO", "Selecting Criteria Values Entered Sucessfully");
    
    var okButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellBtncimgOkBtnPromptsdlg;
    waitForObj(okButton);
    okButton.Click();
    
    var retrievingDataWindow = NameMapping.Sys.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogWaitdlg;
    waitForObj(retrievingDataWindow);
    waitUntilInvisibleOfObj(retrievingDataWindow);
     
    var exportButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelIconimgDhtmllib296;
    waitForObj(exportButton);
    exportButton.Click();
    
    var ExportWindow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cell4;
    waitForObj(ExportWindow);
    ReportUtils.logStep_Screenshot();
        
    var export_OKButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellBtncimgOkBtnIdexportdlg;
    waitForObj(export_OKButton);
    export_OKButton.Click();
    
    aqUtils.Delay(18000, "Waiting to Download Report");
    ReportUtils.logStep("Pass", "Trail Balance Detail Report exported successfully");
    Log.Message("Trail Balance Detail Report exported successfully"); 
    }
  else
   ReportUtils.logStep("Fail", "Selection Criteria Prompt window not displayed");  
 
 var pageName = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.frameIframeleftpanew.cell.panelDivdocname.textContent;

if(pageName.trim() == "Trail Balance Detail" || ImageRepository.Browser_Reporting.TrailBalanceDetail_logo.Exists())
  {
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Trail Balance Detail Screen displayed sucessfully");
     Log.Message("Trail Balance Detail Screen displayed sucessfully");
     } 
  else
     ReportUtils.logStep("Fail", "Trail Balance Detail Screen not displayed");     
     
     
      Sys.Browser("chrome").BrowserWindow(0).Keys("^w");
      var okbutton = Aliases.browser.pageOpendocument.Confirm.Button("OK");
      okbutton.Click();        
}