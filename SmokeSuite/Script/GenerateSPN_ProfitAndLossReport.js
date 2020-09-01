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
function generateProfitAndLoss_Report(){
TextUtils.writeLog("Generation Of Profit And Loss Report"); 

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
gotoProfitAndLossLink();
generateProfitAndLossReport();
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


function gotoProfitAndLossLink()
{
var profitLossLink = Aliases.Maconomy.Reports_GL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McLinkLabelWidget.McTextWidget;
waitForObj(profitLossLink);
Sys.HighlightObject(profitLossLink);
ReportUtils.logStep_Screenshot();
profitLossLink.Click();
ReportUtils.logStep("INFO", "Clicked on Profit And Loss Link");
Log.Message("Clicked on Profit And Loss Link");
}


function generateProfitAndLossReport()
{
var excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "SPN_ProfitAndLossReport";
ExcelUtils.setExcelName(workBook, sheetName, true);


aqUtils.Delay(5000, "Navigating to Browser");
     
  if(ImageRepository.Browser_Reporting.Browser_GLTransaction_Prompt.Exists())
    {
      
    var periodFrom = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cell6;
    periodFrom.Click();
    var fromPeriod =  ExcelUtils.getRowDatas("PeriodFrom",EnvParams.Opco);
    var startDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellLovrightzone2Promptlovzone.textboxPromptlovzoneRightzoneOne;
    waitForObj(startDateField);
    startDateField.setText(fromPeriod);
    
  
    var periodTo = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cell;
    periodTo.Click();
    var toPeriod = ExcelUtils.getRowDatas("PeriodTo",EnvParams.Opco);
    var endDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellLovrightzone2Promptlovzone.textboxPromptlovzoneRightzoneOne;
    waitForObj(endDateField);
    endDateField.setText(toPeriod);
    
    var promptsScroll = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelPromptlist;
    promptsScroll.Click();
    promptsScroll.MouseWheel(-500);

    var exchangeRateDate = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelPromptlist.cell;
    exchangeRateDate.Click();
    var exchangeDate = ExcelUtils.getRowDatas("ExchangeRateDate",EnvParams.Opco);
    var exchangeDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.textboxTextPromptlovzoneRightzon;
    waitForObj(exchangeDateField);
    exchangeDateField.setText(exchangeDate);
    
    var exchangeRateTable = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelPromptlist.cell2;
    exchangeRateTable.Click();
    
    var exchangeRateTableFirstResult = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panel;
    exchangeRateTableFirstResult.Click();
    
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellBtncimgLrzArrowaddPromptlovz;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();
    
    
    var reportCurrency = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelPromptlist.cell3;
    reportCurrency.Click();
    var firstCurrency = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panel;
    firstCurrency.Click();
    
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
    
    aqUtils.Delay(8000, "Waiting to Download Report");
    ReportUtils.logStep("Pass", "Profit And Loss Report exported successfully");
    Log.Message("Profit And Loss Report exported successfully"); 
    }
  else
   ReportUtils.logStep("Fail", "Selection Criteria Prompt window not displayed");  
  
  if(ImageRepository.Browser_Reporting.ProfitAndLoss_logo.Exists())
  {
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Profit And Loss Screen displayed sucessfully");
     Log.Message("ProfitAnd Loss Screen displayed successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Profit And Loss Screen not displayed");            
}