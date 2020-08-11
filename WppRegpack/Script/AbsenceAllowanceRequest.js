//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "AbsenceAllowanceRequest";
var Language = "";
Indicator.Show();
  

ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var AbsenceType,Entrydate,TimeRegistered,Reason1 = ""; 
var Project_manager = "";

//Main Function
function AbsenceAllowanceRequest() {
TextUtils.writeLog("Absence Allowance Request Creation Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "AbsenceAllowanceRequest";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
AbsenceType,Entrydate,TimeRegistered,Reason1 = "";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

try{
getDetails();
goToJobMenuItem();   
AllowanceRquest();   
}
catch(err){
Log.Message(err);
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


//getting data from datasheet
function getDetails(){

Indicator.PushText("Reading Data from Excel");
ExcelUtils.setExcelName(workBook, sheetName, true);
sheetName="AbsenceAllowanceRequest";


AbsenceType = ExcelUtils.getRowDatas("Absence Type",EnvParams.Opco)
Log.Message(AbsenceType)
if((AbsenceType=="")||(AbsenceType==null)){
ValidationUtils.verify(false,true,"Absence Type is Needed to Create a AbsenceAllowanceRequest");
}
    
Entrydate = ExcelUtils.getRowDatas("Entry Date",EnvParams.Opco)
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a AbsenceAllowanceRequest");
}
Log.Message(Entrydate)
  
TimeRegistered = ExcelUtils.getRowDatas("Time Registered",EnvParams.Opco)
Log.Message(TimeRegistered)
if((TimeRegistered=="")||(TimeRegistered==null)){
ValidationUtils.verify(false,true,"Time Registered is Needed to Create a AbsenceAllowanceRequest");
}
  
Reason1 = ExcelUtils.getRowDatas("Reason",EnvParams.Opco)
Log.Message(Reason1)
if((Reason1=="")||(Reason1==null)){
ValidationUtils.verify(false,true,"Reason is Needed to Create a AbsenceAllowanceRequest");
}
  
Indicator.PushText("Playback");

}

// Providing Details for Job Creation
function AllowanceRquest() {
ReportUtils.logStep("INFO", "Enter Job Details");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();

var absenceallowance1 = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(absenceallowance1);
absenceallowance1.HoverMouse();
ReportUtils.logStep_Screenshot("");
absenceallowance1.Click();
aqUtils.Delay(9000, "waiting for new absence allowance");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var newallowancerequestBtn = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.SingleToolItemControl;
                                
WorkspaceUtils.waitForObj(newallowancerequestBtn);
newallowancerequestBtn.HoverMouse();
ReportUtils.logStep_Screenshot("");
newallowancerequestBtn.Click();
TextUtils.writeLog("New Allowance Request is clicked");
aqUtils.Delay(9000, "Checking Labels");

 var C_absencetype = Aliases.Maconomy.AbsenceAllowance1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McValuePickerWidget;
if(AbsenceType!=""){
Sys.HighlightObject(C_absencetype);
C_absencetype.Click();
WorkspaceUtils.SearchByValue(C_absencetype,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Type").OleValue.toString().trim(),AbsenceType,"Absence Type");
TextUtils.writeLog("Entering absence type :"+AbsenceType);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("AllowanceAbsenceType",EnvParams.Opco,"Data Management",AbsenceType)
aqUtils.Delay(6000, "Checking Labels");
}

var c_enterdate=Aliases.Maconomy.AbsenceAllowance1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McDatePickerWidget;
WorkspaceUtils.waitForObj(c_enterdate);
c_enterdate.HoverMouse();
ReportUtils.logStep_Screenshot("");
c_enterdate.Click();
c_enterdate.setText(Entrydate);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("AllowanceAbsenceDate",EnvParams.Opco,"Data Management",Entrydate)
TextUtils.writeLog("Entering EnterDate :"+Entrydate);
aqUtils.Delay(6000, "Checking Labels");


var c_timereg=Aliases.Maconomy.AbsenceAllowance1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McTextWidget;
WorkspaceUtils.waitForObj(c_timereg);
c_timereg.HoverMouse();
ReportUtils.logStep_Screenshot("");
c_enterdate.Click();
c_timereg.setText(TimeRegistered);
TextUtils.writeLog("Entering Time Registered");

var c_reason=Aliases.Maconomy.AbsenceAllowance1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite4.McTextWidget;
WorkspaceUtils.waitForObj(c_reason);
c_timereg.HoverMouse();
ReportUtils.logStep_Screenshot("");
c_reason.Click();
c_reason.setText(Reason1+" "+STIME);
var Type = Reason1+" "+STIME;      
TextUtils.writeLog("Entering Reason");
aqUtils.Delay(6000, "Checking Labels");

var create=Aliases.Maconomy.AbsenceAllowance1.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
WorkspaceUtils.waitForObj(create);
create.HoverMouse();
ReportUtils.logStep_Screenshot("");
create.Click();
aqUtils.Delay(6000, "Checking Labels");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var all = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim())
WorkspaceUtils.waitForObj(all);
all.HoverMouse();
ReportUtils.logStep_Screenshot("");
all.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var table = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
waitForObj(table)
Sys.HighlightObject(table);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var c_enterdate=Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(c_enterdate);
c_enterdate.HoverMouse();
ReportUtils.logStep_Screenshot("");
c_enterdate.Click();
c_enterdate.setText(Entrydate);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
c_enterdate.Keys("[Tab][Tab][Tab]")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var Reson=Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget3
WorkspaceUtils.waitForObj(Reson);
Reson.HoverMouse();
ReportUtils.logStep_Screenshot("");
Reson.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Reson.setText(Type);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000, "Checking Labels");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000, "Checking Labels");
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  Log.Message(table.getItem(v).getText_2(3).OleValue.toString().trim())
  Log.Message(Type)
  Log.Message(table.getItem(v).getText_2(3).OleValue.toString().trim()==Type)
  if(table.getItem(v).getText_2(3).OleValue.toString().trim()==Type){ 
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
 ValidationUtils.verify(flag,true,"Created Allowance Request is available in system");
 
var submit = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.SingleToolItemControl2;
WorkspaceUtils.waitForObj(submit);
submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
submit.Click();
aqUtils.Delay(6000, "Checking Labels");
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Allowance Reason",EnvParams.Opco,"Data Management",Type)
var supervisorinfo = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
WorkspaceUtils.waitForObj(supervisorinfo);
supervisorinfo.HoverMouse();
supervisorinfo.Click();
var mess= supervisorinfo.getText();
Log.Message("Approver for Allowance"+mess);
ReportUtils.logStep_Screenshot("");
TextUtils.writeLog("Supervisor Information :"+mess);
aqUtils.Delay(6000, "Checking Labels")
}

//Go To Job from Menu
function goToJobMenuItem(){


var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.TimeExpense.Exists()){
 ImageRepository.ImageSet.TimeExpense.Click();// GL
}
else if(ImageRepository.ImageSet.TimeExpense1.Exists()){
ImageRepository.ImageSet.TimeExpense1.Click();
}
else{
ImageRepository.ImageSet.TimeExpense2.Click();
}


var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//Delay(3000);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
}
} 

ReportUtils.logStep("INFO", "Moved to Absence from Absence Menu");
TextUtils.writeLog("Entering into Absence from Time & Expense Menu");
}


