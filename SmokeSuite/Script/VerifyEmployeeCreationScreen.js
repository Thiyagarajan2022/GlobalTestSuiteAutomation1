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
function verifyNewEmployeeCreation(){
TextUtils.writeLog("Verification Of Employee Creation"); 

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
verifyEmployeeCreation();
WorkspaceUtils.closeAllWorkspaces();
}
  catch(err){
    Log.Message(err);
  }
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.HumanResources.Exists()){
ImageRepository.ImageSet.HumanResources.Click();
}
else{
     ReportUtils.logStep("Fail", "Human Resources Workspace not displayed");
     Log.Message("Human Resources Workspace not displayed");
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
report_mngmnt.ClickItem("|Employees");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Employees");
}
}
aqUtils.Delay(2000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Employees Management workspace from Human Resources Menu");
Log.Message("Entering Employees workspace from Human Resources Menu");
}
}

function verifyEmployeeCreation(){ 

var empTab= Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(empTab);

Sys.HighlightObject(empTab);


var newEmpButton = Aliases.Maconomy.HumanResources.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;

Sys.HighlightObject(newEmpButton);
newEmpButton.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}


var newEmpInfoScreen = Aliases.Maconomy.HR_Emp_NewEmployeeScreen.Composite.Label;
WorkspaceUtils.waitForObj(newEmpInfoScreen);

if(newEmpInfoScreen.isVisible())
{
     ReportUtils.logStep_Screenshot();
     Sys.HighlightObject(newEmpInfoScreen);
     ReportUtils.logStep("Pass", "New Employee Screen loaded successfully");
     Log.Message("New Employee Screen loaded successfully");
     } 
  else
     ReportUtils.logStep("Fail", "New Employee Screen not loaded");


var newEmpInfoCancelBtn = Aliases.Maconomy.HR_Emp_NewEmployeeScreen.Composite.Composite.Composite.Composite.Button;
WorkspaceUtils.waitForObj(newEmpInfoCancelBtn);

newEmpInfoCancelBtn.Click();

if(ImageRepository.ImageSet.LoadedBox.Exists()){}
aqUtils.Delay(2000, "Waiting to Load");


}


