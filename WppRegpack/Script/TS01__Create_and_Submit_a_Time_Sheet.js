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
 * This script create and submit Timesheet
 * @author  : Muthu Kumar M
 * @version : 3.0
 * Created Date :02/10/2021
 * Modified Date(MM/DD/YYYY) : 12/22/2021
*/

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateTimesheet";
var Language = "";
Indicator.Show();
Indicator.PushText("waiing for window to open");
ExcelUtils.setExcelName(workBook, sheetName, true);


var z=0;
var invisible_Startindex;
var invisible_Endindex = 6;
var visible_Startindex = 0;
var visible_Endindex;
var jobNumber,weekno,workID,Descrip,mon,tue,wed,thu,fri,sat,sun,EmployeeNumber,EmployeeName,startdate,enddate = "";
var mon_Des,tue_Des,wed_Des,thu_Des,fri_Des,sat_Des,sun_Des = "";
var numberOfHours =0;
var Maconomy_ParentAddress,Maconomy_Index = "";

function CreateTimeSheet(){ 
TextUtils.writeLog("Timesheet Creation Started"); 
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



//aqTestCase.Begin("Job Creation", "zfj://CH1-30");
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateTimesheet";


ExcelUtils.setExcelName(workBook, sheetName, true);
z=0;
invisible_Startindex = "";
invisible_Endindex = 6;
visible_Startindex = 0;
visible_Endindex = "";
jobNumber,weekno,workID,Descrip,mon,tue,wed,thu,fri,sat,sun,EmployeeNumber,EmployeeName,startdate,enddate = "";
mon_Des,tue_Des,wed_Des,thu_Des,fri_Des,sat_Des,sun_Des = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create Timesheet started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME);
try{
getDetails();
goTo_TimeSheet();
  
sheetName = "CreateTimesheet";
checking_Week_inCalender();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

WorkspaceUtils.closeAllWorkspaces();
}catch(err){ 
  Log.Message(err);
}

//aqTestCase.End();
}


//getting data from datasheet
function getDetails(){
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
  }
  
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Timesheet");


ExcelUtils.setExcelName(workBook, sheetName, true);
weekno = ExcelUtils.getRowDatas("Weekno",EnvParams.Opco)
if((weekno==null)||(weekno=="")){ 
ValidationUtils.verify(false,true,"Weekno is Needed to Create Timesheet");
}
 for(var i=1;i<=10;i++){
sheetName = "JobBudgetCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
 workID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
 Log.Message(workID);
if((workID!="")&&(workID!=null)&&(workID.indexOf("T")!=-1)){
 break;
}else{ 
 workID = ""; 
}
}
sheetName = "CreateTimesheet";
ExcelUtils.setExcelName(workBook, sheetName, true);
if((workID==null)||(workID=="")){ 
workID = ExcelUtils.getRowDatas("Workcode",EnvParams.Opco)
}
if((workID==null)||(workID=="")){ 
ValidationUtils.verify(false,true,"Workcode is Needed to Create Timesheet");
}
Descrip = ExcelUtils.getRowDatas("Description",EnvParams.Opco)
if((Descrip==null)||(Descrip=="")){ 
ValidationUtils.verify(false,true,"Description is Needed to Create Timesheet");
}
mon = ExcelUtils.getRowDatas("Mon",EnvParams.Opco)
//if((mon==null)||(mon=="")){ 
//ValidationUtils.verify(false,true,"Time for Monday is Needed to Create Timesheet");
//}
tue = ExcelUtils.getRowDatas("Tue",EnvParams.Opco)
//if((tue==null)||(tue=="")){ 
//ValidationUtils.verify(false,true,"Time for Tuesday is Needed to Create Timesheet");
//}
wed = ExcelUtils.getRowDatas("Wed",EnvParams.Opco)
//if((wed==null)||(wed=="")){ 
//ValidationUtils.verify(false,true,"Time for Wednessday is Needed to Create Timesheet");
//}
thu = ExcelUtils.getRowDatas("Thu",EnvParams.Opco)
//if((thu==null)||(thu=="")){ 
//ValidationUtils.verify(false,true,"Time for Thursday is Needed to Create Timesheet");
//}
fri= ExcelUtils.getRowDatas("Fri",EnvParams.Opco)
//if((fri==null)||(fri=="")){ 
//ValidationUtils.verify(false,true,"Time for Friday is Needed to Create Timesheet");
//}
sat = ExcelUtils.getRowDatas("Sat",EnvParams.Opco)
sun = ExcelUtils.getRowDatas("Sun",EnvParams.Opco)

mon_Des = ExcelUtils.getRowDatas("Mon Description",EnvParams.Opco)

tue_Des = ExcelUtils.getRowDatas("Tue Description",EnvParams.Opco)

wed_Des = ExcelUtils.getRowDatas("Wed Description",EnvParams.Opco)

thu_Des = ExcelUtils.getRowDatas("Thu Description",EnvParams.Opco)

fri_Des = ExcelUtils.getRowDatas("Fri Description",EnvParams.Opco)

sat_Des = ExcelUtils.getRowDatas("Sat Description",EnvParams.Opco)

sun_Des= ExcelUtils.getRowDatas("Sun Description",EnvParams.Opco)


}


// Navigating to Time & Expenses from Time & Expenses Menu
function goTo_TimeSheet(){

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_timesheet_from_workspace(); //Select Timesheet & Expense Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());

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
  
  var status = true;
  while(status)// If Year is need to check code need to start from here
  if(week1.getText()==weekno){ 
    status = false;
    selectDay(1);
    break;
  }else if(week2.getText()==weekno){
    status = false;
    selectDay(8);
    break;
  }else if(week3.getText()==weekno){
    status = false;
    selectDay(15);
    break;
  }else if(week4.getText()==weekno){
    status = false;
    selectDay(22);
    break;
  }else if(week5.getText()==weekno){
    status = false;
    selectDay(29);
    break;
  }else if(week6.getText()==weekno){
    status = false;
    selectDay(36);
    break;
  }else{
    
  if(week6.getText()>weekno){ 
      previousMonth.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      previousMonth.Click();
      aqUtils.Delay(4000, "Changing Month");
  }
  if(week6.getText()<weekno){ 
      nextMonth.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      nextMonth.Click();
      aqUtils.Delay(4000, "Changing Month");
  }
  
  }

}



function selectDay(startday){
weekseparate = false;
var days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday);  
   for(i=1;i<7;i++){ 
   var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i).getText();
   if(day=="1"){
    invisible_Startindex = i;
    visible_Endindex = i-1;
   weekseparate = true;
   break;
   
}
}

if(!weekseparate){ 
var Rejectline = false;
  for(i=0;i<7;i++){

   var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i);
   if((day.background==8689648)||(day.background==16777215)||(day.background==3405809)){
   if(day.background==3405809){
   Rejectline = true;
   ReportUtils.logStep("INFO","Selected week is already Rejected, Altering Infomation in Registration panel")
   }
     }else{ 
       ValidationUtils.verify(false,true,"Value is already Submitted or Approved");
     }
     }
days.HoverMouse();
ReportUtils.logStep_Screenshot("");  
days.Click();   
TextUtils.writeLog("Week has been Selected");

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
WorkspaceUtils.waitForObj(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4));


var create_Timesheet = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim());
if(create_Timesheet!=null)
//var create_Timesheet = Aliases.Maconomy.CreateTimesheet.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl.getText();
if(create_Timesheet.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim()){
Sys.HighlightObject(create_Timesheet);
var create_Timesheet = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim());
create_Timesheet.HoverMouse();
ReportUtils.logStep_Screenshot(""); 
create_Timesheet.Click();
}
TextUtils.writeLog("Entering Timesheet Details");

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
 
    var linegrid = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2)
    var linecount = linegrid.getItemCount()
    var linecount_Status = true;
    for(var gg=3;gg<11;gg++){  
    if((linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()=="")||(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()==null))
    linecount_Status = false;    
    }
    if(!Rejectline){  
    var addline = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Add Time Sheet Line (Ctrl+M)").OleValue.toString().trim());
addline.HoverMouse();
ReportUtils.logStep_Screenshot("");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    addline.Click();
    }
    else{
    var Registration = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    Registration.Click();
    }

waitUntil_MaconomyScreen_loaded_Completely();

    var keep = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McPlainCheckboxView",1)
    WorkspaceUtils.waitForObj(keep)
    keep.Keys("[Tab][Tab][Tab]");
    aqUtils.Delay(1000, Indicator.Text);

    var Job = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McValuePickerWidget",2,"McGrid")
    WorkspaceUtils.waitForObj(Job)
    Job.Click();
    if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_1(Job,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number")
    Job
    }
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

    var work = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McValuePickerWidget",2,"McGrid")
    work.Click();
    if(work.getText()!=workID){
    WorkspaceUtils.SearchByValue(work,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),workID,"WorkCode")
    work
    }
    aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var description = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McTextWidget",3,"McGrid")

    if(((Descrip!="")||(Descrip!=null))&&(description.getText()!=Descrip)){
    description.setText(Descrip);
    ValidationUtils.verify(true,true,"Description is Entered");
    }
    var Time_for_day = switchcase(0,6)
    

    for(var kk=0;kk<7;kk++){
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var days = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McTextWidget",3,"McGrid");

if(((switchDayValue(Time_for_day[kk])!="")||(switchDayValue(Time_for_day[kk])!=null))&&(days.getText()!=switchDayValue(Time_for_day[kk]))){
days.Keys(switchDayValue(Time_for_day[kk]));
numberOfHours = numberOfHours + parseInt(switchDayValue(Time_for_day[kk]));
ValidationUtils.verify(true,true,"Time is Entered for "+Time_for_day[kk]);
}
    z++;
  
     
}

aqUtils.Delay(3000, "Timesheet is Saved");
Log.Message("Total hours:"+numberOfHours);    

    var Save;
    Save = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl", "Save Time Sheet Line");
    Log.Message(Save.FullName);
    
    if(Save.isEnabled_2){
      Save.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      Save.Click();
      TextUtils.writeLog("Timesheet is Saved");
      aqUtils.Delay(3000, "Timesheet is Saved");

    ValidationUtils.verify(true,true,"Line of Timesheet is Saved");
    if(eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "*").WndCaption ==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses - Registrations").OleValue.toString().trim()){ 
      var Ok = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses - Registrations").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Ok.Click();
      aqUtils.Delay(3000, "Timesheet is Saved");
    }
    
waitUntil_MaconomyScreen_loaded_Completely();
  
  var balance = getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget",1,2);
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array("McTextWidget", "9", "true");
  obj = balance.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Parent.ChildCount ==4){
    balance = obj[i_count];
    break;      
  }
}
Sys.HighlightObject(balance);

  if(balance.getText().OleValue.toString().trim().indexOf("-")!=-1){
  ValidationUtils.verify(false,true,"Balance is not Zero or Timesheet is not balanced");
  }else{ 
  ValidationUtils.verify(true,true,"Timesheet is balanced");
  TextUtils.writeLog("Timesheet is balanced");
  }
  ImageRepository.ImageSet.Maximize1.Click();
  waitUntil_MaconomyScreen_loaded_Completely();
  
  var Description = "";
  
  Description = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Descriptions")

  WorkspaceUtils.waitForObj(Description);
  Description.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
//  16777215
//11862010 - for Yellow
Log.Message(switchDayValue("Mon"))
Log.Message(switchDayValue("Tue"))
Log.Message(switchDayValue("Wed"))
Log.Message(switchDayValue("Thu"))
Log.Message(switchDayValue("Fri"))
Log.Message(switchDayValue("Sat"))
Log.Message(switchDayValue("Sun"))
aqUtils.Delay("1000","Loading");
  var monday_Des = "";



monday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if((switchDayValue("Mon")!="")&&(switchDayValue("Mon")!=null))
  if(monday_Des.background=="11862010"){
    if((mon_Des!="")&&(mon_Des!=null))
    monday_Des.setText(mon_Des);
    else
    ValidationUtils.verify(true,false,"Description for Monday is required")
  }else{ 
    if((mon_Des!="")&&(mon_Des!=null))
    monday_Des.setText(mon_Des);
  }
  
    var tuesday_Des = "";
    tuesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if((switchDayValue("Tue")!="")&&(switchDayValue("Tue")!=null))
  if(tuesday_Des.background=="11862010"){
    if((tue_Des!="")&&(tue_Des!=null))
    tuesday_Des.setText(tue_Des);
    else
    ValidationUtils.verify(true,false,"Description for Tuesday is required")
  }else{ 
    if((tue_Des!="")&&(tue_Des!=null))
    tuesday_Des.setText(tue_Des);
  } 
    var wednesday_Des = "";
    wednesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
    

  if((switchDayValue("Wed")!="")&&(switchDayValue("Wed")!=null))
  if(wednesday_Des.background=="11862010"){
    if((wed_Des!="")&&(wed_Des!=null))
    wednesday_Des.setText(wed_Des);
    else
    ValidationUtils.verify(true,false,"Description for Wednessday is required")
  }else{ 
    if((wed_Des!="")&&(wed_Des!=null))
    wednesday_Des.setText(wed_Des);
  }
    var thursday_Des = "";
  thursday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if((switchDayValue("Thu")!="")&&(switchDayValue("Thu")!=null))
  if(thursday_Des.background=="11862010"){
    if((thu_Des!="")&&(thu_Des!=null))
    thursday_Des.setText(thu_Des);
    else
    ValidationUtils.verify(true,false,"Description for Thursday is required")
  }else{ 
    if((thu_Des!="")&&(thu_Des!=null))
    thursday_Des.setText(thu_Des);
  }
    var friday_Des = "";
    friday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if((switchDayValue("Fri")!="")&&(switchDayValue("Fri")!=null))
  if(friday_Des.background=="11862010"){
    if((fri_Des!="")&&(fri_Des!=null))
    friday_Des.setText(fri_Des);
    else
    ValidationUtils.verify(true,false,"Description for Friday is required")
  }else{ 
    if((fri_Des!="")&&(fri_Des!=null))
    friday_Des.setText(fri_Des);
  }
    var saturday_Des = ""
    saturday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if((switchDayValue("Sat")!="")&&(switchDayValue("Sat")!=null))
  if(saturday_Des.background=="11862010"){
    if((sat_Des!="")&&(sat_Des!=null))
    saturday_Des.setText(sat_Des);
    else
    ValidationUtils.verify(true,false,"Description for Saturday is required")
  }else{ 
    if((sat_Des!="")&&(sat_Des!=null))
    saturday_Des.setText(sat_Des);
  }
    var sunday_Des = "";
    sunday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if((switchDayValue("Sun")!="")&&(switchDayValue("Sun")!=null))
  if(sunday_Des.background=="11862010"){
    if((sun_Des!="")&&(sun_Des!=null))
    sunday_Des.setText(sun_Des);
    else
    ValidationUtils.verify(true,false,"Description for Sunday is required")
  }else{ 
    if((sun_Des!="")&&(sun_Des!=null))
    sunday_Des.setText(sun_Des);
  }
  
        var Save;
        Save = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl", "Save Time Sheet Line (Enter)");

  WorkspaceUtils.waitForObj(Save)
  ReportUtils.logStep_Screenshot("");
  Save.Click();
  aqUtils.Delay(1000, "Description is Saved");
waitUntil_MaconomyScreen_loaded_Completely();

  var close = "";

if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible()){
   close = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
  }else{
   close = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "")
  }
  WorkspaceUtils.waitForObj(close)
  ReportUtils.logStep_Screenshot("");
  close.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//===============================================================
      var Submit = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").isVisible()){
    Submit = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    }else{
    Submit = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    }
    Sys.HighlightObject(Submit);
    Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Submit.Click();
    ValidationUtils.verify(true,true,"Timesheet is Submit");
    TextUtils.writeLog("Timesheet is Submit");
    
waitUntil_MaconomyScreen_loaded_Completely();
    

      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management",EmployeeNumber)
      ExcelUtils.WriteExcelSheet("Timesheet Employee Name",EnvParams.Opco,"Data Management",EmployeeName)
      ExcelUtils.WriteExcelSheet("Timesheet Week No",EnvParams.Opco,"Data Management",weekno)
      ExcelUtils.WriteExcelSheet("Timesheet Job No",EnvParams.Opco,"Data Management",jobNumber)
      ExcelUtils.WriteExcelSheet("Timesheet Hours",EnvParams.Opco,"Data Management",numberOfHours)
      
      TextUtils.writeLog("Timesheet is submitted by Employee No:"+EmployeeNumber);
      TextUtils.writeLog("Timesheet is submitted by Employee Name:"+EmployeeName);
      TextUtils.writeLog("Timesheet is Created for Week No:"+weekno);


    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }

}
}
else{ 
var Rejectline = false;
  for(i=0;i<7;i++){


   var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i);
   if((day.background==8689648)||(day.background==16777215)||(day.background==3405809)){
   if(day.background==3405809){
   Rejectline = true;
   ReportUtils.logStep("INFO","Selected week is already Rejected, Altering Infomation in Registration panel")
   }
     }else{ 
       ValidationUtils.verify(false,true,"Value is already Submitted or Approved");
     }
     } 

    var day = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+visible_Startindex);
   day.HoverMouse();
ReportUtils.logStep_Screenshot("");
    day.Click();
    TextUtils.writeLog("Week has been Selected");
    aqUtils.Delay(2000, Indicator.Text);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
var create_Timesheet = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim());


if(create_Timesheet!=null)
if(create_Timesheet.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim()){
Sys.HighlightObject(create_Timesheet);
var create_Timesheet = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim());
   create_Timesheet.HoverMouse();
ReportUtils.logStep_Screenshot("");
create_Timesheet.Click();
}
TextUtils.writeLog("Entering Timesheet Details");
    aqUtils.Delay(2000, "Entering Timesheet Details");
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }

waitUntil_MaconomyScreen_loaded_Completely();


    var linegrid = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);



    var linecount = linegrid.getItemCount();
var linecount_Status = true;
    for(var gg=3;gg<11;gg++){  

    if((linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()=="")||(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()==null))
    linecount_Status = false;    
    }
    if(!Rejectline){   
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

    var addline = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);


    addline.HoverMouse();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
ReportUtils.logStep_Screenshot("");
    addline.Click();
    }
    else{
    var Registration = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    Registration.Click();
    }
    
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

    var keep = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "");


    WorkspaceUtils.waitForObj(keep);
    keep.Keys("[Tab][Tab][Tab]");
    aqUtils.Delay(1000, Indicator.Text);

if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var Job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McValuePickerWidget", "", 2,60000);
    else
    var Job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2)
    WorkspaceUtils.waitForObj(Job);
    Job.Click();
    if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_1(Job,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number");
    Job
    }

    aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

    if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var work = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
    else
    var work = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
    work.Click();
    if(work.getText()!=workID){
    WorkspaceUtils.SearchByValue(work,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),workID,"WorkCode");
    work
    }

aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);

if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var description = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    else
    var description = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    if(((Descrip!="")||(Descrip!=null))&&(description.getText()!=Descrip)){
    description.setText(Descrip);
    ValidationUtils.verify(true,true,"Description is Entered");
    }
    var Time_for_day = switchcase(visible_Startindex,visible_Endindex)
    var z=0;
    for(var k=0;k<7;k++){ 
      if((k>=visible_Startindex)&&(k<=visible_Endindex)){
    aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);

        if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    else
    var days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
if(((switchDayValue(Time_for_day[z])!="")||(switchDayValue(Time_for_day[z])!=null))&&(days.getText()!=switchDayValue(Time_for_day[z]))){
days.Keys(switchDayValue(Time_for_day[z]));
ValidationUtils.verify(true,true,"Time is Entered for "+Time_for_day[z]);
}
    z++;
      }
      else{ 
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09); 
      }
    }
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    var Save;
    Save = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl", "Save Time Sheet Line");

    Log.Message(Save.FullName);
    if(Save.isEnabled_2){
Save.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Save.Click();
    aqUtils.Delay(3000, "Saving Timesheet");
    ValidationUtils.verify(true,true,"Line of Timesheet is Saved");
    TextUtils.writeLog("Timesheet is Saved");
    if(eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "*").WndCaption ==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses - Registrations").OleValue.toString().trim()){ 
      var Ok = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses - Registrations").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Ok.Click();
      aqUtils.Delay(3000, "Saving Timesheet");
    }
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  var balance = getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget",1,2);
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array("McTextWidget", "9", "true");
  obj = balance.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Parent.ChildCount ==4){
    balance = obj[i_count];
    break;      
  }
}
Sys.HighlightObject(balance);

  if(balance.getText().OleValue.toString().trim().indexOf("-")!=-1){
  ValidationUtils.verify(false,true,"Balance is not Zero or Timesheet is not balanced");
  }else{ 
  ValidationUtils.verify(true,true,"Timesheet is balanced");
  TextUtils.writeLog("Timesheet is balanced");
  }
  //===============================
  var Description = "";
  Description = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Descriptions")

  WorkspaceUtils.waitForObj(Description);
  Description.Click();
    ImageRepository.ImageSet.Maximize1.Click();

//  16777215
//11862010 - for Yellow
Log.Message(switchDayValue("Mon"))
Log.Message(switchDayValue("Tue"))
Log.Message(switchDayValue("Wed"))
Log.Message(switchDayValue("Thu"))
Log.Message(switchDayValue("Fri"))
Log.Message(switchDayValue("Sat"))
Log.Message(switchDayValue("Sun"))

  if(0<=visible_Endindex)
  if((switchDayValue("Mon")!="")&&(switchDayValue("Mon")!=null)){   
var monday_Des = "";
monday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

if(monday_Des.background=="11862010"){
    if((mon_Des!="")&&(mon_Des!=null))
    monday_Des.setText(mon_Des);
    else
    ValidationUtils.verify(true,false,"Description for Monday is required")
  }else{ 
    
var monday_Des = "";
monday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  
    if((mon_Des!="")&&(mon_Des!=null))
    monday_Des.setText(mon_Des);
  }
 } 

  if(1<=visible_Endindex)
  if((switchDayValue("Tue")!="")&&(switchDayValue("Tue")!=null)){

  var tuesday_Des = "";
  tuesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(tuesday_Des.background=="11862010"){
    if((tue_Des!="")&&(tue_Des!=null))
    tuesday_Des.setText(tue_Des);
    else
    ValidationUtils.verify(true,false,"Description for Tuesday is required")
  }else{ 
    
  var tuesday_Des = "";
  tuesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  
    if((tue_Des!="")&&(tue_Des!=null))
    tuesday_Des.setText(tue_Des);
  } 
  }
  

  if(2<=visible_Endindex)
  if((switchDayValue("Wed")!="")&&(switchDayValue("Wed")!=null)){

  var wednesday_Des = "";
  wednesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  

  if(wednesday_Des.background=="11862010"){
    if((wed_Des!="")&&(wed_Des!=null))
    wednesday_Des.setText(wed_Des);
    else
    ValidationUtils.verify(true,false,"Description for Wednessday is required")
  }else{ 
    
  var wednesday_Des = "";
  wednesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  
  
    if((wed_Des!="")&&(wed_Des!=null))
    wednesday_Des.setText(wed_Des);
  }
  }
  

  if(3<=visible_Endindex)
  if((switchDayValue("Thu")!="")&&(switchDayValue("Thu")!=null)){

  
  var thursday_Des = "";
  thursday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(thursday_Des.background=="11862010"){
    if((thu_Des!="")&&(thu_Des!=null))
    thursday_Des.setText(thu_Des);
    else
    ValidationUtils.verify(true,false,"Description for Thursday is required")
  }else{ 
    
  var thursday_Des = "";
  thursday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  
  
    if((thu_Des!="")&&(thu_Des!=null))
    thursday_Des.setText(thu_Des);
  }
 } 

  if(4<=visible_Endindex)
  if((switchDayValue("Fri")!="")&&(switchDayValue("Fri")!=null)){

    
  var friday_Des = "";
  friday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  if(friday_Des.background=="11862010"){
    if((fri_Des!="")&&(fri_Des!=null))
    friday_Des.setText(fri_Des);
    else
    ValidationUtils.verify(true,false,"Description for Friday is required")
  }else{ 
    
  var friday_Des = "";
  friday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  
    if((fri_Des!="")&&(fri_Des!=null))
    friday_Des.setText(fri_Des);
  }
  }

  if(5<=visible_Endindex)
  if((switchDayValue("Sat")!="")&&(switchDayValue("Sat")!=null)){

    
  var saturday_Des = ""
  saturday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  

  if(saturday_Des.background=="11862010"){
    if((sat_Des!="")&&(sat_Des!=null))
    saturday_Des.setText(sat_Des);
    else
    ValidationUtils.verify(true,false,"Description for Saturday is required")
  }else{ 
    
  var saturday_Des = ""
  saturday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  

  
    if((sat_Des!="")&&(sat_Des!=null))
    saturday_Des.setText(sat_Des);
  }
  }

  if(6<=visible_Endindex)
  if((switchDayValue("Sun")!="")&&(switchDayValue("Sun")!=null)){

    
  var sunday_Des = "";
  sunday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  
   
  if(sunday_Des.background=="11862010"){
    if((sun_Des!="")&&(sun_Des!=null))
    sunday_Des.setText(sun_Des);
    else
    ValidationUtils.verify(true,false,"Description for Sunday is required")
  }else{ 
    
  var sunday_Des = "";
  sunday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  

    if((sun_Des!="")&&(sun_Des!=null))
    sunday_Des.setText(sun_Des);
  }
  }
  
      var Save;
      

   if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).isVisible()){
Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  }
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).isVisible()){
Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);  
}

  WorkspaceUtils.waitForObj(Save)
  ReportUtils.logStep_Screenshot("");
  Save.Click();
  aqUtils.Delay(1000, "Description is Saved");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var close = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible()){
   close = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
  }else{
   close = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "")
  }
  WorkspaceUtils.waitForObj(close)
  ReportUtils.logStep_Screenshot("");
  close.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  //================================
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var Submit = "";
    
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").isVisible()){
    Submit = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    }else{
    Submit = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    }
    
    Sys.HighlightObject(Submit);
Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Submit.Click();
    ValidationUtils.verify(true,true,"Timesheet is Submit");
    TextUtils.writeLog("Timesheet is Submit");
    }
    aqUtils.Delay(4000, "Submitting Timesheet");
    
waitUntil_MaconomyScreen_loaded_Completely();
    
    
var nextMonth = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").isVisible()){
    nextMonth = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 3);
    }else{
    nextMonth = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 3)
    }
nextMonth.HoverMouse();
ReportUtils.logStep_Screenshot("");
    nextMonth.Click();
    TextUtils.writeLog("Navigating to next month to create Time sheet for that week");
    aqUtils.Delay(3000, "Moving to next month");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var days = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").isVisible()){
    days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 1+invisible_Startindex);  
    }
else{ 
   days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 1+invisible_Startindex);    
}
    days.HoverMouse();
ReportUtils.logStep_Screenshot("");
    days.Click();
    TextUtils.writeLog("Week has been Selected");
    aqUtils.Delay(2000, "Week has been Selected");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var create_Timesheet = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim());
if(create_Timesheet!=null)
if(create_Timesheet.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim()){
Sys.HighlightObject(create_Timesheet);
  var create_Timesheet = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Time Sheet").OleValue.toString().trim());
create_Timesheet.HoverMouse();
ReportUtils.logStep_Screenshot("");
create_Timesheet.Click();
}
TextUtils.writeLog("Entering Timesheet Details");
    aqUtils.Delay(3000, "Entering Timesheet Details");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var linegrid = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).isVisible()){
linegrid = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
}else{
linegrid = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
}
    var linecount = linegrid.getItemCount()
var linecount_Status = true;
    for(var gg=3;gg<11;gg++){  
    if((linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()!="")||(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()!=null))
    linecount_Status = false;    
    }
    if(!Rejectline){   
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var addline = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).isVisible()){
addline = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4)
}else{
addline = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
}
    addline.HoverMouse();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
ReportUtils.logStep_Screenshot("");
    addline.Click();
    }
    else{
    var Registration = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    Registration.Click();
    }

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

    if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var keep = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "");
    else
    var keep = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "")
    
    WorkspaceUtils.waitForObj(keep)
    keep.Keys("[Tab][Tab][Tab]");
    aqUtils.Delay(1000, Indicator.Text);

    if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var Job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McValuePickerWidget", "", 2,60000);
    else
    var Job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2)
    WorkspaceUtils.waitForObj(Job)
    Job.Click();
if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_1(Job,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number");
    Job
    }

    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

    if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var work = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
    else
    var work = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
    work.Click();
    if(work.getText()!=workID){
    WorkspaceUtils.SearchByValue(work,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),workID,"WorkCode");
    work
    }

aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var description = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    else
    var description = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    if(((Descrip!="")||(Descrip!=null))&&(description.getText()!=Descrip)){
    description.setText(Descrip);
    ValidationUtils.verify(true,true,"Description is Entered");
    }
    var Time_for_day = switchcase(invisible_Startindex,invisible_Endindex)
    var z=0;
    
    for(var k=0;k<7;k++){ 
      if((k>=invisible_Startindex)&&(k<=invisible_Endindex)){ 
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    var days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    else
    var days = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
if(((switchDayValue(Time_for_day[z])!="")||(switchDayValue(Time_for_day[z])!=null))&&(days.getText()!=switchDayValue(Time_for_day[z]))){
days.Keys(switchDayValue(Time_for_day[z]));
ValidationUtils.verify(true,true,"Time is Entered for "+Time_for_day[z]);
}
    z++;
      }
      else{
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09); 
      }
    }

 
    var Save;
    Save = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl", "Save Time Sheet Line");

    if(Save.isEnabled_2){
Save.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Save.Click();
    aqUtils.Delay(3000, "Saving Timesheet");
    ValidationUtils.verify(true,true,"Line of Timesheet is Saved");
    TextUtils.writeLog("Timesheet is Saved");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    if(eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "*").WndCaption ==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses - Registrations").OleValue.toString().trim()){ 
      var Ok = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses - Registrations").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Ok.Click();
      aqUtils.Delay(3000, "Timesheet is Saved");
    }

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  var balance = getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,"McGroupWidget",1,2);
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array("McTextWidget", "9", "true");
  obj = balance.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Parent.ChildCount ==4){
    balance = obj[i_count];
    break;      
  }
}
Sys.HighlightObject(balance);

  if(balance.getText().OleValue.toString().trim().indexOf("-")!=-1){
  ValidationUtils.verify(false,true,"Balance is not Zero or Timesheet is not balanced");
  }else{ 
  ValidationUtils.verify(true,true,"Timesheet is balanced");
  TextUtils.writeLog("Timesheet is balanced");
  }
    
      //===============================
    var Description = "";
    Description = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Descriptions")
  WorkspaceUtils.waitForObj(Description);
  Description.Click();
    ImageRepository.ImageSet.Maximize1.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

//  16777215
//11862010 - for Yellow
Log.Message(switchDayValue("Mon"))
Log.Message(switchDayValue("Tue"))
Log.Message(switchDayValue("Wed"))
Log.Message(switchDayValue("Thu"))
Log.Message(switchDayValue("Fri"))
Log.Message(switchDayValue("Sat"))
Log.Message(switchDayValue("Sun"))


  if((0>=invisible_Startindex)&&(0<=invisible_Endindex))
  if((switchDayValue("Mon")!="")&&(switchDayValue("Mon")!=null)){

var monday_Des = "";
monday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(monday_Des.background=="11862010"){
    if((mon_Des!="")&&(mon_Des!=null))
    monday_Des.setText(mon_Des);
    else
    ValidationUtils.verify(true,false,"Description for Monday is required")
  }else{ 
var monday_Des = "";
monday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


    if((mon_Des!="")&&(mon_Des!=null))
    monday_Des.setText(mon_Des);
  }
  }
  

  if((1>=invisible_Startindex)&&(1<=invisible_Endindex))
  if((switchDayValue("Tue")!="")&&(switchDayValue("Tue")!=null)){

var tuesday_Des = "";
tuesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(tuesday_Des.background=="11862010"){
    if((tue_Des!="")&&(tue_Des!=null))
    tuesday_Des.setText(tue_Des);
    else
    ValidationUtils.verify(true,false,"Description for Tuesday is required")
  }else{ 
    
var tuesday_Des = "";
tuesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


    if((tue_Des!="")&&(tue_Des!=null))
    tuesday_Des.setText(tue_Des);
  }
  } 
  

  if((2>=invisible_Startindex)&&(2<=invisible_Endindex))
  if((switchDayValue("Wed")!="")&&(switchDayValue("Wed")!=null)){

var wednesday_Des = "";
wednesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(wednesday_Des.background=="11862010"){
    if((wed_Des!="")&&(wed_Des!=null))
    wednesday_Des.setText(wed_Des);
    else
    ValidationUtils.verify(true,false,"Description for Wednessday is required")
  }else{ 
    
var wednesday_Des = "";
wednesday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


    if((wed_Des!="")&&(wed_Des!=null))
    wednesday_Des.setText(wed_Des);
  }
  }
  

  if((3>=invisible_Startindex)&&(3<=invisible_Endindex))
  if((switchDayValue("Thu")!="")&&(switchDayValue("Thu")!=null)){

    
var thursday_Des = "";
thursday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(thursday_Des.background=="11862010"){
    if((thu_Des!="")&&(thu_Des!=null))
    thursday_Des.setText(thu_Des);
    else
    ValidationUtils.verify(true,false,"Description for Thursday is required")
  }else{ 
    
var thursday_Des = "";
thursday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


    if((thu_Des!="")&&(thu_Des!=null))
    thursday_Des.setText(thu_Des);
  }
  }
  

  if((4>=invisible_Startindex)&&(4<=invisible_Endindex))
  if((switchDayValue("Fri")!="")&&(switchDayValue("Fri")!=null)){

    
var friday_Des = "";
friday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(friday_Des.background=="11862010"){
    if((fri_Des!="")&&(fri_Des!=null))
    friday_Des.setText(fri_Des);
    else
    ValidationUtils.verify(true,false,"Description for Friday is required")
  }else{ 
    
var friday_Des = "";
friday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


    if((fri_Des!="")&&(fri_Des!=null))
    friday_Des.setText(fri_Des);
  }
  }
  

  if((5>=invisible_Startindex)&&(5<=invisible_Endindex))
  if((switchDayValue("Sat")!="")&&(switchDayValue("Sat")!=null)){

    
var saturday_Des = "";
saturday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(saturday_Des.background=="11862010"){
    if((sat_Des!="")&&(sat_Des!=null))
    saturday_Des.setText(sat_Des);
    else
    ValidationUtils.verify(true,false,"Description for Saturday is required")
  }else{ 
    
var saturday_Des = "";
saturday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


    if((sat_Des!="")&&(sat_Des!=null))
    saturday_Des.setText(sat_Des);
  }
  }
  

  if((6>=invisible_Startindex)&&(6<=invisible_Endindex))
  if((switchDayValue("Sun")!="")&&(switchDayValue("Sun")!=null)){

    
var sunday_Des = "";
sunday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);


  if(sunday_Des.background=="11862010"){
    if((sun_Des!="")&&(sun_Des!=null))
    sunday_Des.setText(sun_Des);
    else
    ValidationUtils.verify(true,false,"Description for Sunday is required")
  }else{ 
    
var sunday_Des = "";
sunday_Des = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGroupWidget",1).SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);

  
    if((sun_Des!="")&&(sun_Des!=null))
    sunday_Des.setText(sun_Des);
  }
  }
  

      var Save;

   if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).isVisible()){
Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  }
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).isVisible()){
Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);  
}
  WorkspaceUtils.waitForObj(Save)
  ReportUtils.logStep_Screenshot("");
  Save.Click();
  aqUtils.Delay(1000, "Description is Saved");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var close = "";
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible()){
   close = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
  }else{
   close = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "")
  }
  WorkspaceUtils.waitForObj(close)
  ReportUtils.logStep_Screenshot("");
  close.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  //================================
      var Submit = "";
      Submit = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl","Submit");

      Sys.HighlightObject(Submit);
      Submit.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      Submit.Click();
      ValidationUtils.verify(true,true,"Timesheet is Submit");
      TextUtils.writeLog("Timesheet is Submit");
   }
   
   
waitUntil_MaconomyScreen_loaded_Completely();
  
     
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management",EmployeeNumber)
      ExcelUtils.WriteExcelSheet("Timesheet Employee Name",EnvParams.Opco,"Data Management",EmployeeName)
      ExcelUtils.WriteExcelSheet("Timesheet Week No",EnvParams.Opco,"Data Management",weekno)
      ExcelUtils.WriteExcelSheet("Timesheet Job No",EnvParams.Opco,"Data Management",jobNumber)
      ExcelUtils.WriteExcelSheet("Timesheet Hours",EnvParams.Opco,"Data Management",numberOfHours)
      
      TextUtils.writeLog("Timesheet is submitted by Employee No:"+EmployeeNumber);
      TextUtils.writeLog("Timesheet is submitted by Employee Name:"+EmployeeName);
      TextUtils.writeLog("Timesheet is Created for Week No:"+weekno);
      
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
}
}






























function switchcase(start,end){ 
var array = [];
var temp = "";
var j=0;
for(var i=start;i<=end;i++){
  switch (i) {
     case 0:{
     temp = "Mon";
     }
     break;
     case 1:{
     temp = "Tue";
     }
     break;
     case 2:{
     temp = "Wed";
     }
     break;
     case 3:{
     temp = "Thu";
     }
     break;
     case 4:{
     temp = "Fri";
     }
     break;
     case 5:{
     temp = "Sat";
     }
     break;
     case 6:{
     temp = "Sun";
     }
     break;
  }
  array[j]=temp;
  j++;
  }
  return array;
}

function switchDayValue(day){ 
var array = [];
var temp = "";

  switch (day) {
     case "Mon":{
     temp = mon;
     }
     break;
     case "Tue":{
     temp = tue;
     }
     break;
     case "Wed":{
     temp = wed;
     }
     break;
     case "Thu":{
     temp = thu;
     }
     break;
     case "Fri":{
     temp = fri;
     }
     break;
     case "Sat":{
     temp = sat;
     }
     break;
     case "Sun":{
     temp = sun;
     }
     break;
  }
 

  return temp;
}


function readlog(){

 

sheetName = "JobCreation";

ExcelUtils.setExcelName(workBook, sheetName, true);

comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)

if((comapany==null)||(comapany=="")){

ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");

}

Job_group = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)

if((Job_group==null)||(Job_group=="")){

ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");

}

var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)

var name =  LogReport_name(ExlArray,comapany,Job_group);

var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";

//Log.Message("Notepad :"+notepadPath)

return TextUtils.readDetails(notepadPath,"Job Number");

//Log.Message( readDetails("C:\\Users\\674087\\Documents\\TestComplete 14 Projects\\After Stuart Discussion\\WppRegression_v12.50\\WppRegPack\\RegressionLogs\\TESTAPAC\\Regression\\China\\1307-Sudler China(MDS)_Client Billable.txt","Job Number") );

}

 

 

function LogReport_name(ExcelData,value,JG){

var compStatus = "";

      for(var exl =0;exl<ExcelData.length;exl++){

        var splits = [];

        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));

        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);

      if(splits[0]==value.toString().trim()){

        compStatus = ExcelData[exl]+"_"+JG;

        break;

      }

      }

//Log.Message(compStatus);

return compStatus

}

 

 

function getExcelData_Company(rowidentifier,column) {

excelData =[]; 
var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
var id =0;
var colsList = [];
var temp ="";
     while (!DDT.CurrentDriver.EOF()) {

       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){

        try{

         temp = temp+xlDriver.Value(column).toString().trim();

         }

        catch(e){

        temp = "";

        }

      break;

      }

 

    xlDriver.Next();

     }

    

     if(temp.indexOf("*")!=-1){

     var excelData =  temp.split("*");

     

     }else if(temp.length>0){

      excelData[0] = temp;

     }

    

     DDT.CloseDriver(xlDriver.Name);

for(var i=0;i<excelData.length;i++)

// Log.Message(excelData[i]);

     return excelData;

 

}



