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
function verifyTimeExpense(){
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
gotoTimeTab();
verifyWeekScreen();
verifyExpenseScreen();
gotoTimeSheetLookUp();
verifyTimeSheetsLookup();
gotoAbsenceWorkspace();
verifyAbsenceWorkspace();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Time_Expense.Exists()){
ImageRepository.ImageSet.Time_Expense.Click();
}
else{
     ReportUtils.logStep("Fail", "Time_Expense  Section not displayed");
     Log.Message("Time_Expense Section not displayed");
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
report_mngmnt.ClickItem("|Time & Expenses");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Time & Expenses");
}

} 
ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
Log.Message("Entering into Time & Expenses from Time & Expenses Menu");
}

function gotoTimeTab(){ 
aqUtils.Delay(2000, "Waiting to Load");

if(ImageRepository.ImageSet.LoadedBox.Exists())
  {   
  var timeTab =Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(timeTab);
  timeTab.Click();
  ReportUtils.logStep("INFO", "Time Tab");
  Log.Message("Clicked Time Tab")
  }
}

function verifyWeekScreen()
{
aqUtils.Delay(2000, "Waiting to Load");
var weekSection = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(weekSection);
ReportUtils.logStep_Screenshot();
weekSection.Click();
ReportUtils.logStep("INFO", "Clicked on Week Section under Time");
Log.Message("Clicked on Week Section under Time");

var calendar = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McCalendarWidget.McDateChooser;

if(calendar.isVisible())
{
     Sys.HighlightObject(calendar);
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Time Sheet creation workspace is loaded successfully");
     Log.Message("Time Sheet creation workspace is loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Time Sheet creation workspace is not loaded");             
}

function verifyExpenseScreen()
{
aqUtils.Delay(2000, "Waiting to Load");

var expenseTab = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl
Sys.HighlightObject(expenseTab);
ReportUtils.logStep_Screenshot();
expenseTab.Click();

aqUtils.Delay(2000, "Waiting to Load");

var expenseSheet = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(expenseSheet);

ReportUtils.logStep_Screenshot();

var newExpensebutton = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;

if(newExpensebutton.isVisible())
{
     Sys.HighlightObject(newExpensebutton);
     ReportUtils.logStep("Pass", "Expense Sheet creation workspace is loaded successfully");
     Log.Message("Expense Sheet creation workspace is loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Expense Sheet creation workspace is not loaded");             
}

function gotoTimeSheetLookUp(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Time_Expense.Exists()){
ImageRepository.ImageSet.Time_Expense.Click();
}
else{
     ReportUtils.logStep("Fail", "Time_Expense  Section not displayed");
     Log.Message("Time_Expense Section not displayed");
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
report_mngmnt.ClickItem("|Time Sheet Lookups");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Time Sheet Lookups");
} 
}
aqUtils.Delay(2000, "Waiting to Load"); 
if(ImageRepository.ImageSet.ExcelImage.Exists())
  {
  ReportUtils.logStep("INFO", "Moved to Time Sheet Lookups from Time & Expenses Menu");
  Log.Message("Entering into Time Sheet Lookups from Time & Expenses Menu");
  }
}

function verifyTimeSheetsLookup()
{
aqUtils.Delay(6000, "Waiting to Load");
//var listTimeSheetsTab = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
//Sys.HighlightObject(listTimeSheetsTab);
//ReportUtils.logStep_Screenshot();
//listTimeSheetsTab.Click();
//ReportUtils.logStep("INFO", "Clicked on List of TimeSheetsTab under Time Sheets Lookup");
//Log.Message("Clicked on List of TimeSheetsTab under Time Sheets Lookup");

var timeSheetsLinesTab = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TimeSheetLines
Sys.HighlightObject(timeSheetsLinesTab);

if(timeSheetsLinesTab.isVisible())
{
     Sys.HighlightObject(timeSheetsLinesTab);
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Time Sheet Lookups workspace is loaded successfully");
     Log.Message("Time Sheet Lookups workspace is loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Time Sheet Lookups workspace is not loaded");             
}

function gotoAbsenceWorkspace(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Time_Expense.Exists()){
ImageRepository.ImageSet.Time_Expense.Click();
}
else{
     ReportUtils.logStep("Fail", "Time_Expense  Section not displayed");
     Log.Message("Time_Expense Section not displayed");
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
report_mngmnt.ClickItem("|Absence");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Absence");
} 
} 
aqUtils.Delay(2000, "Waiting to Load");

if(ImageRepository.ImageSet.Legend.Exists())
  {
  ReportUtils.logStep("INFO", "Moved to Absence from Time & Expenses Menu");
  Log.Message("Entering into Absence from Time & Expenses Menu");
  }
}

function verifyAbsenceWorkspace()
{
aqUtils.Delay(4000, "Waiting to Load");
var absenceInfo = Aliases.Maconomy.TimeExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget;

  if(absenceInfo.isVisible())
  {
     Sys.HighlightObject(absenceInfo);
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "Absence workspace is loaded successfully");
     Log.Message("Absence workspace is loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "Absence workspace is not loaded");  
                
}



