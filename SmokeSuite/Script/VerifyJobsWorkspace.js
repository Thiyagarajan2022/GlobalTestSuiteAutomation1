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
function verifyJobs(){
TextUtils.writeLog("Verification Of Jobs Worksapce"); 

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
verifyJobsList();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.JobsMenu.Exists())
ImageRepository.ImageSet.JobsMenu.Click();
else if (ImageRepository.ImageSet.Jobs_Workspace.Exists())
ImageRepository.ImageSet.Jobs_Workspace.Click();
else{
     ReportUtils.logStep("Fail", "Jobs Workspace not displayed");
     Log.Message("Jobs Workspace not displayed");
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
report_mngmnt.ClickItem("|Jobs");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Jobs");
}
}
aqUtils.Delay(3000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Jobs Workspace from Jobs Menu");
Log.Message("Entering Jobs Workspace from Jobs Menu");
}
}

function verifyJobsList(){ 
aqUtils.Delay(5000, "Waiting to Load");
var jobListTab = Aliases.Maconomy.JobsWorkspace.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(jobListTab);

var allMyJobs = Aliases.Maconomy.JobsWorkspace.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
Sys.HighlightObject(allMyJobs);
allMyJobs.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }

var internalJobs = Aliases.Maconomy.JobsWorkspace.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button2;
Sys.HighlightObject(internalJobs);
internalJobs.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }

var billableJobs = Aliases.Maconomy.JobsWorkspace.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button3

Sys.HighlightObject(billableJobs);
billableJobs.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }

ReportUtils.logStep("INFO", "Toogling Job Filters");
ReportUtils.logStep_Screenshot(); 
  
 var companyNoSearch =Aliases.Maconomy.JobsWorkspace.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
 companyNoSearch.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
 
 var curr = Aliases.Maconomy.JobsWorkspace.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McPopupPickerWidget2;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 curr.Click();
 aqUtils.Delay(3000, "Waiting to Load currerncies");
 WorkspaceUtils.DropDownList("GBP","Currency");
 
if(curr.isVisible())
{
     ReportUtils.logStep_Screenshot();
     ReportUtils.logStep("Pass", "List of Jobs are loaded successfully");
     Log.Message("List of Jobs are loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "List of Jobs are not loaded");
}


