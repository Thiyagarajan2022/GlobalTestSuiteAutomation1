﻿//USEUNIT EnvParams
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
function generateBalanceSheet_Report(){
TextUtils.writeLog("Verification Of Balance Sheet Report"); 

//Setting Language in WorkspaceUtils
Language = "";
Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

//Checking Login
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4).Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}


//Initializing Variables

try{
//Entering Report Management
WorkspaceUtils.closeAllWorkspaces();
gotoMenu(); 
gotoGLScreen();
gotoBalanceSheetLink();
generateBalanceSheetReport();
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
ReportUtils.logStep("INFO", "Moved to GL Reports from Reporting Menu");
Log.Message("Entering into GL Reports from Reporting Menu");
}

function gotoGLScreen(){ 
var glScreen = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(glScreen);
Sys.HighlightObject(glScreen);
glScreen.Click();
ReportUtils.logStep("INFO", "Clicked GL Menu Section");
Log.Message("Clicked GL Menu Section")
}

function gotoBalanceSheetLink()
{
var balanceSheetLink = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McLinkLabelWidget.McTextWidget;
waitForObj(balanceSheetLink);
Sys.HighlightObject(balanceSheetLink);
ReportUtils.logStep_Screenshot();
balanceSheetLink.Click();
ReportUtils.logStep("INFO", "Clicked on Balance Sheet Link");
Log.Message("Clicked on Balance Sheet Link");
}


function generateBalanceSheetReport()
{
  
var excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "BalanceSheetReport";
ExcelUtils.setExcelName(workBook, sheetName, true);
var company =  ExcelUtils.getRowDatas("Company",EnvParams.Opco);


aqUtils.Delay(5000, "Navigating to Browser");

aqUtils.Delay(8000, "Navigating to Browser and waiting for Data Protection window");
 if(ImageRepository.Browser_Reporting.DataProtectionDialog.Exists())
  {
    var ok_button = Aliases.browser.pageOpendocument.buttonOk
    ok_button.Click();
    }
  else{
  var dataProtectionDialog = Aliases.browser.pageOpendocument.panel
  waitForObj(dataProtectionDialog);
  var ok_button = Aliases.browser.pageOpendocument.buttonOk
  ok_button.Click();
  }  

aqUtils.Delay(5000, "Waiting for Prompt Dialog");    
    var promptDialog = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogPromptsdlg.table.cell.table.cell
    waitForObj(promptDialog);
    
    var companySearch = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.textboxLovwidgetpromptlovzoneSea;
    waitForObj(companySearch);
    companySearch.SetText(company);
    
    var searchIcon =  Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellIconmidIconmenuIconLovwidget.panel
    waitForObj(searchIcon);
    searchIcon.Click();
    
    var searchResult = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.panelMlstBodylovwidgetpromptlovz.table.cell.panelMclm0;
    
    Log.Message("Company is Listed in Search Result")
    searchResult.Click();
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.buttonRealbtnLrzArrowaddPromptlo;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();
    
    var periodFrom = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogPromptsdlg.table.panelTreecontPromptlist.tableCwpromptspromptlisttwe1.cell;
    periodFrom.Click();
    var fromPeriod =  ExcelUtils.getRowDatas("PeriodFrom",EnvParams.Opco);
    var startDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.tableLovwidgetpromptlovzone.cell.panelDivInputLovwidgetpromptlovz.textboxLovwidgetpromptlovzoneTex;
    waitForObj(startDateField);
    startDateField.setText(fromPeriod);
    
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.buttonRealbtnLrzArrowaddPromptlo;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();
    
    
    var periodTo = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogPromptsdlg.table.panelTreecontPromptlist.tableCwpromptspromptlisttwe2.cell;
    periodTo.Click();
    var toPeriod = ExcelUtils.getRowDatas("PeriodTo",EnvParams.Opco);
    var endDateField = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.tableLovwidgetpromptlovzone.cell.panelDivInputLovwidgetpromptlovz.textboxLovwidgetpromptlovzoneTex;
    waitForObj(endDateField);
    endDateField.setText(toPeriod);
    
    var moveRightArrow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.buttonRealbtnLrzArrowaddPromptlo;
    waitForObj(moveRightArrow);
    moveRightArrow.Click();

    ReportUtils.logStep("INFO", "Selecting Criteria Values Entered Sucessfully");
    
    var okButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.tableOkBtnPromptsdlg.cellBtncimgOkBtnPromptsdlg;
    waitForObj(okButton);
    okButton.Click();
    
    var retrievingDataWindow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogWaitdlg;
    waitForObj(retrievingDataWindow);
    waitUntilInvisibleOfObj(retrievingDataWindow);
     
    if(EnvParams.instanceData.includes("APAC"))
    {
    var exportButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellIconmidDhtmllib307.panel.panelIconimgDhtmllib307;
    waitForObj(exportButton);
    exportButton.Click();
    
    var ExportWindow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogIdexportdlg.table.cell.table.cell;
    waitForObj(ExportWindow);
    ReportUtils.logStep_Screenshot();
        
    var export_OKButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogIdexportdlg.table.tableOkBtnIdexportdlg.cellBtncimgOkBtnIdexportdlg;
    waitForObj(export_OKButton);
    export_OKButton.Click();
    }
    else if(EnvParams.instanceData.includes("EMEA"))
    {
    var exportButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellIconmidDhtmllib292.panel.panelIconimgDhtmllib292;
    waitForObj(exportButton);
    exportButton.Click();
    
    var ExportWindow = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogIdexportdlg.table.cell.table.cell
    waitForObj(ExportWindow);
    ReportUtils.logStep_Screenshot();
        
    var export_OKButton = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.cellTdDialogIdexportdlg.table.tableOkBtnIdexportdlg.cellBtncimgOkBtnIdexportdlg;
    waitForObj(export_OKButton);
    export_OKButton.Click();  
    }
    aqUtils.Delay(25000, "Waiting to Download Report");
    ReportUtils.logStep("Pass", "Balance Sheet Report exported successfully");
    Log.Message("GL Balance Sheet Report exported successfully"); 
   
  
aqUtils.Delay(2000, "Loading Balance Sheet Screen");
  var pageName = "";
   if(EnvParams.instanceData.includes("EMEA"))
    pageName = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.frameIframeleftpanew.textnode.textContent;
   else if(EnvParams.instanceData.includes("APAC")) 
    pageName = Aliases.browser.pageOpendocument.frameOpendocchildframe.frameWebiviewframe.frameIframeleftpanew.panelDivdocname.textContent;
  
  if(pageName.trim()== "Balance Sheet" || ImageRepository.Browser_Reporting.BalanceSheet_Logo.Exists())
  {
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Balance Sheet Screen displayed sucessfully");
     Log.Message("Balance Sheet Screen displayed sucessfully");
     } 
  else
     ReportUtils.logStep("Fail", "Balance Sheet Screen not displayed");  
      var brow = Sys.Browser("chrome"); brow.Close(); if(brow.Exists) brow.Terminate();         
}