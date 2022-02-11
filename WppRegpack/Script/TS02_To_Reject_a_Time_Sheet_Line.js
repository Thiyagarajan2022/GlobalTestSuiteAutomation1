﻿//USEUNIT ActionUtils
//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ObjectUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/** 
 * This script reject the created timesheet
 * @author  : Muthu Kumar M
 * @version : 3.0
 * Modified Date(MM/DD/YYYY) : 01/06/2022
 */
 

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "RejectTimesheet";
var Language = "";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");
  
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var EmpNumber,weekno,TYear,comapany,level = "";
var Approve_Level = [];
var ApproveInfo = [];
var level = 0; 
var Project_manager = "";
var Maconomy_ParentAddress,Maconomy_Index = "";

// Main Function
function RejectTimesheet(){ 
TextUtils.writeLog("Rejecting Timesheet Started"); 
Indicator.PushText("waiting for window to reponse");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Create Budget
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);



//Re-Intialize Variable
EmpNumber,weekno,TYear,comapany = "";
level = 0;
Approve_Level = [];
ApproveInfo = [];
comapany = EnvParams.Opco;
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "RejectTimesheet";


ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Rejecting Timesheet started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME);

try{
getDetails();
ExcelUtils.setExcelName(workBook, sheetName, true);  
goTo_TimeSheet();  

sheetName = "RejectTimesheet";
checking_Week_inCalender();
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);

// Rejecting Timesheet
var i=0
level = i;
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
Workspace_Client.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);


Maconomy_Index = WorkspaceUtils.Maconomy_Parent;
Maconomy_Index++;
WorkspaceUtils.Maconomy_Parent = Maconomy_Index;
Log.Message(Maconomy_Index);

// Restarting maconomy with Approver Logins
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
var Macscreen = WorkspaceUtils.switch_Maconomy(temp[2])
if(Macscreen=="Screen Not Found"){
Restart.login(temp[2]);
}
aqUtils.Delay(5000, Indicator.Text);

Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(temp[2]);
//Refreshing To-Do's List to find Submitted Job
aqUtils.Delay(5000, Indicator.Text);

ActionUtils.ToDos_Selection(Maconomy_ParentAddress, level, temp[3], "Approve Time Sheet Line", null, "Approve Time Sheet Line (Substitute)", null)
RejectTimesht(temp[0],temp[1],temp[2]);

WorkspaceUtils.closeAllWorkspaces();
}
catch(err){ 
  Log.Message(err);
}
}


function getDetails(){ 
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  EmpNumber = ReadExcelSheet("Timesheet Employee Name",EnvParams.Opco,"Data Management");
  if((EmpNumber=="")||(EmpNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  EmpNumber = ExcelUtils.getRowDatas("Employee Name",EnvParams.Opco)
  }  
  if((EmpNumber=="")||(EmpNumber==null))
  ValidationUtils.verify(false,true,"Employee Name is needed to Approve Timesheet");
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  weekno = ReadExcelSheet("Timesheet Week No",EnvParams.Opco,"Data Management");
  if((weekno=="")||(weekno==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  weekno = ExcelUtils.getRowDatas("Weekno",EnvParams.Opco)
  }  
  if((weekno=="")||(weekno==null))
  ValidationUtils.verify(false,true,"Week No is needed to Approve Timesheet");
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  TYear = ReadExcelSheet("Timesheet Year",EnvParams.Opco,"Data Management");
  if((TYear=="")||(TYear==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  TYear = ExcelUtils.getRowDatas("Year",EnvParams.Opco)
  }  
  if((TYear=="")||(TYear==null))
  ValidationUtils.verify(false,true,"Year is needed to Approve Timesheet"); 
}


// Navigating to Time & Expenses from Time & Expenses Menu
function goTo_TimeSheet(){

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_timesheet_from_workspace(); //Select Timesheet & Expense Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());

}


function RejectTimesht(comID,EmpName,AprName){ 
//  aqUtils.Delay(5000, Indicator.Text);
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists())  { 
  
}

if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);
} 

            
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "");
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2); 
var firstCell = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
//aqUtils.Delay(2000, Indicator.Text);
firstCell.Keys("[Tab]");
var EName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
var week = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
var Yer = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
WorkspaceUtils.waitForObj(EName);
EName.Keys(EmpName);
aqUtils.Delay(1000, Indicator.Text);
EName.Keys("[Tab][Tab][Tab]");
WorkspaceUtils.waitForObj(week);
week.Keys(weekno);
aqUtils.Delay(1000, Indicator.Text);
week.Click();
week.Keys("[Tab][Tab]");
Yer.Keys(TYear);
var closefilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
for(var v=0;v<table.getItemCount();v++){ 
WorkspaceUtils.waitForObj(table);
var flag=false;
  if((table.getItem(v).getText_2(1).OleValue.toString().trim()==EmpName)||
  (table.getItem(v).getText_2(4).OleValue.toString().trim().indexOf(weekno)!=-1) ||
  (table.getItem(v).getText_2(6).OleValue.toString().trim()==TYear)){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true; 
    table.Keys("[Down]");
    ValidationUtils.verify(flag,true,"Created Timesheet is available in system");
    ReportUtils.logStep_Screenshot();  
    if(flag){ 
   closefilter.Click();
//aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var approveButton = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 5);
var rejectButton = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);
WorkspaceUtils.waitForObj(rejectButton);
rejectButton.HoverMouse();
ReportUtils.logStep_Screenshot();
rejectButton.Click();
aqUtils.Delay(15000, Indicator.Text);
var remarks = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
remarks.Keys("Rejected");
//aqUtils.Delay(5000, Indicator.Text);
var Rejectall = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject all").OleValue.toString().trim());
Rejectall.HoverMouse();
ReportUtils.logStep_Screenshot();
Rejectall.Click();
//ValidationUtils.verify(true,true,"Created Timesheet is Rejected by :"+AprName)
//TextUtils.writeLog("Created Timesheet is Rejected by :"+AprName);
//aqUtils.Delay(5000, Indicator.Text);


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var ApvPerson = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
WorkspaceUtils.waitForObj(ApvPerson);
ApvPerson.Click();
var loginPer = eval(Maconomy_ParentAddress).WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while ((ApvPerson.getText().OleValue.toString().trim().indexOf("Rejected")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)&&(i!=600))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}
//Log.Message(ApvPerson.getText().OleValue.toString().trim())
//Log.Message(loginPer)
//Log.Message((ApvPerson.getText().OleValue.toString().trim().indexOf("Rejected")!=-1))
//Log.Message((ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1))
  if(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1){
  ValidationUtils.verify(true,true,"Created Timesheet is Rejected by :"+loginPer)
  TextUtils.writeLog("Created Timesheet is Rejected by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Created Timesheet is Rejected by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Created Timesheet is Rejected by :"+loginPer+ "But its Not Reflected")
  }
  


var approve_Bar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 3).WaitSWTObject("TabControl", "",1,60000);
WorkspaceUtils.waitForObj(approve_Bar);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
//Delay(2000);
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
//Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
var All_approver = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 5,60000);
WorkspaceUtils.waitForObj(All_approver);
All_approver.Click();
//Delay(3000);

ReportUtils.logStep_Screenshot();
//var approver_table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
}
var info_Bar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(info_Bar);
info_Bar.Click();
//Delay(4000);

ImageRepository.ImageSet.Forward.Click();
//Delay(4000);
}
var showFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2)
WorkspaceUtils.waitForObj(showFilter);
showFilter.HoverMouse();
ReportUtils.logStep_Screenshot();
showFilter.Click();
aqUtils.Delay(2000, "searching data in tables");  
   }
    
//    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
 
aqUtils.Delay(1000, "Waiting to Undo");

for(var v=0;v<table.getItemCount();v++){
  table.Keys("[Up]");
  }
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
for(var v=0;v<table.getItemCount();v++){ 
WorkspaceUtils.waitForObj(table);
var flag=false;
  if((table.getItem(v).getText_2(1).OleValue.toString().trim()==EmpName)||
  (table.getItem(v).getText_2(4).OleValue.toString().trim().indexOf(weekno)!=-1) ||
  (table.getItem(v).getText_2(6).OleValue.toString().trim()==TYear)){ 
    flag=true; 
    table.Keys("[Down]");
    ReportUtils.logStep_Screenshot();  
    if(flag){ 
   closefilter.Click();
//aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var undo = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
WorkspaceUtils.waitForObj(undo);
undo.HoverMouse();
//ReportUtils.logStep_Screenshot();
undo.Click();
aqUtils.Delay(4000, "Undo All Approvals/Rejected");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var showFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2)
WorkspaceUtils.waitForObj(showFilter);
showFilter.HoverMouse();
ReportUtils.logStep_Screenshot();
showFilter.Click();
aqUtils.Delay(2000, "searching data in tables"); 
   }
 
//    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}


}




function checking_Week_inCalender(){ 
  
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
Log.Message(Maconomy_ParentAddress)

  var EmployeeNumber = eval(getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McValuePickerWidget",3));
  EmployeeNumber = EmployeeNumber.getText().OleValue.toString().trim();


  //Wait till Employee Name and Number be visible to proceed further
  var Visiblestatus = true;
  while(Visiblestatus){ 
    EmployeeNumber = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McValuePickerWidget",3).getText().OleValue.toString().trim();
    EmployeeName = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McValuePickerWidget",2).getText().OleValue.toString().trim();
    if((EmployeeNumber!="")&&(EmployeeNumber!=null)){ 
      aqUtils.Delay(2000, Indicator.Text);
      Visiblestatus = false;
    }
  }
  
  
  var previousMonth = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGenericButton",2);
  var nextMonth = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGenericButton",3);
  var week1 = getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"DateChooser$CellLabel",1,6);
  var week2 = getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"DateChooser$CellLabel",2,6);
  var week3 = getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"DateChooser$CellLabel",3,6);
  var week4 = getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"DateChooser$CellLabel",4,6);
  var week5 = getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"DateChooser$CellLabel",5,6);
  var week6 = getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"DateChooser$CellLabel",6,6);
  var YearMonth = getObjectAddress_withSingleProperty_Check(Maconomy_ParentAddress,"Label")
  var previousYear = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGenericButton",1);
  var nextYear = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGenericButton",4);
  var YearMon = YearMonth.getText().OleValue.toString().trim();
  
  var Year = YearMon.split(" ");
  var status = true;
  while(status){// If Year is need to check code need to start from here
  
  if(Year[1]>TYear){ 
previousYear.HoverMouse();
ReportUtils.logStep_Screenshot("");
  previousYear.Click();
  aqUtils.Delay(4000, "Changing Year");
  YearMon = YearMonth.getText().OleValue.toString().trim();
  Year = YearMon.split(" ");
  if(Year[1]==TYear){ 
    break;
  }
  }
  
  
  if(Year[1]<TYear){ 
nextYear.HoverMouse();
ReportUtils.logStep_Screenshot("");
  nextYear.Click();
  aqUtils.Delay(4000, "Changing Year");
  YearMon = YearMonth.getText().OleValue.toString().trim();
  Year = YearMon.split(" ");
  if(Year[1]==TYear){ 
    break;
  }
  }
  if(Year[1]==TYear){ 
  break;
  }
  }
  
 var status = true;
  while(status){
  if(week1.getText().OleValue.toString()==weekno){ 
    status = false;
    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 1);
//    day.Click();
    selectDay(1);
    break;
  }else if(week2.getText().OleValue.toString()==weekno){
    status = false;
    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 8);
//    day.Click();
    selectDay(8);
    break;
  }else if(week3.getText().OleValue.toString()==weekno){
    status = false;
    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 15);
//    day.Click();
    selectDay(15);
    break;
  }else if(week4.getText().OleValue.toString()==weekno){
    status = false;
    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 22);
//    day.Click();
    selectDay(22);
    break;
  }else if(week5.getText().OleValue.toString()==weekno){
    status = false;
    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 29);
//    day.Click();
    selectDay(29);
    break;
  }else if(week6.getText().OleValue.toString()==weekno){
    status = false;
    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 36);
//    day.Click();
    selectDay(36);
    break;
  }else{
//  Log.Message("week6 :"+week6.getText());
//  Log.Message("weekno :"+weekno); 
  if(week6.getText()>weekno){ 
previousMonth.HoverMouse();
ReportUtils.logStep_Screenshot("");
  previousMonth.Click();
  aqUtils.Delay(4000, "Changing Month");
//  Delay(5000);  
  }
  if(week6.getText()<weekno){ 
nextMonth.HoverMouse();
ReportUtils.logStep_Screenshot("");
  nextMonth.Click();
  aqUtils.Delay(4000, "Changing Month");
//  Delay(5000);  
  }
  }
  }
  
}

function selectDay(startday){
var days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday);

days.HoverMouse();
ReportUtils.logStep_Screenshot("");  
days.Click();  
//aqUtils.Delay(3000, Indicator.Text);
TextUtils.writeLog("Week has been Selected to get Approver");
var approve_Bar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(approve_Bar);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
//Delay(2000);
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
//Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
var All_approver = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 5,60000);
WorkspaceUtils.waitForObj(All_approver);
//Delay(1000);
All_approver.Click();
//Delay(3000);
ReportUtils.logStep_Screenshot();
var approver_table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
WorkspaceUtils.waitForObj(approver_table);
//  Log.Message(approver_table.getItemCount());
var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   approvers="";
   if(approver_table.getItem(z).getText_2(8)!="Approved"){
   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   ReportUtils.logStep("INFO","Employee Approver level : " +z+ " Approver :" +approvers);
//   Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
   Approve_Level[y] = comapany+"*"+EmpNumber+"*"+approvers;
   Log.Message(Approve_Level[y])
   y++;
   }
}
}
TextUtils.writeLog("Finding approvers for Created Timesheet");
var info_Bar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
WorkspaceUtils.waitForObj(info_Bar)
info_Bar.Click();
//Delay(4000);

ImageRepository.ImageSet.Forward.Click();
//Delay(4000);
}
WorkspaceUtils.closeAllWorkspaces();
}







