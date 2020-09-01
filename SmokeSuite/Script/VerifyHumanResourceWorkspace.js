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
function verifyHumanResource(){
TextUtils.writeLog("Verification Of Human Resource Worksapce"); 

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
verifyEmployeesWorkspace();
gotoUsersWorkspace();
verifyUsersWorkspace();
WorkspaceUtils.closeAllWorkspaces();
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

if(ImageRepository.ImageSet.HumanResources.Exists()){
ImageRepository.ImageSet.HumanResources.Click();
}
else{
     ReportUtils.logStep("Fail", "Human Resources Workspace not displayed");
     Log.Message("Human Resources Workspace not displayed");
}

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|Employees");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Employees");
}
}
aqUtils.Delay(3000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Employees Workspace from Human Resources Menu");
Log.Message("Entering Employees Workspace from Human Resources Menu");
}
}

function verifyEmployeesWorkspace(){ 
aqUtils.Delay(5000, "Waiting to Load");
var employeesTab= Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(employeesTab);

var currentEmp = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
Sys.HighlightObject(currentEmp);
currentEmp.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }
ReportUtils.logStep("INFO", "Toogling Employee Filters");
ReportUtils.logStep_Screenshot();
  
var allEmp = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button2
Sys.HighlightObject(allEmp);
allEmp.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }
 
var nowShowingResults = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.Label;

if(nowShowingResults.isVisible())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(nowShowingResults);
     ReportUtils.logStep("Pass", "List of Employees are loaded successfully");
     Log.Message("List of Employees are loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "List of Employees are not loaded");
}

function gotoUsersWorkspace(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.HumanResources.Exists()){
ImageRepository.ImageSet.HumanResources.Click();
}
else{
     ReportUtils.logStep("Fail", "Human Resources  Section not displayed");
     Log.Message("Human Resources Section not displayed");
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
report_mngmnt.ClickItem("|Users");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Users");
} 
} 
aqUtils.Delay(2000, "Waiting to Load");

if(ImageRepository.ImageSet.ExcelImage.Exists())
  {
  ReportUtils.logStep("INFO", "Moved to Users from HumanResources Menu");
  Log.Message("Entering into Users from HumanResources Menu");
  }
}

function verifyUsersWorkspace(){ 
aqUtils.Delay(5000, "Waiting to Load");
var usersTab= Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite3.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(usersTab);

var activeUsers = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button3;
Sys.HighlightObject(activeUsers);
activeUsers.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }
ReportUtils.logStep("INFO", "Toogling Users Filters");
ReportUtils.logStep_Screenshot();
  
var blockedUsers = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button4;
Sys.HighlightObject(blockedUsers);
blockedUsers.Click();
 if(ImageRepository.ImageSet.LoadedBox.Exists()){    
  }
 
var nowShowingResults = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.Label;

if(nowShowingResults.isVisible())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(nowShowingResults);
     ReportUtils.logStep("Pass", "List of Employees are loaded successfully");
     Log.Message("List of Employees are loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "List of Employees are not loaded");
}
