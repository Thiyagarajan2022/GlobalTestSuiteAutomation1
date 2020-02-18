//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart 
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
//var jobNumber = "1307200330";
//var weekno = "40"
//var workID = "T1001"
//var Descrip = "Billable - Timesheet"
//var mon = "7.5"
//var tue = "7.5"
//var wed = "7.5"
//var thu = "7.5"
//var fri = "7.5"
//var sat,sun = "";

//getting data from datasheet
function getDetails(){
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
if((mon==null)||(mon=="")){ 
ValidationUtils.verify(false,true,"Time for Monday is Needed to Create Timesheet");
}
tue = ExcelUtils.getRowDatas("Tue",EnvParams.Opco)
if((tue==null)||(tue=="")){ 
ValidationUtils.verify(false,true,"Time for Tuesday is Needed to Create Timesheet");
}
wed = ExcelUtils.getRowDatas("Wed",EnvParams.Opco)
if((wed==null)||(wed=="")){ 
ValidationUtils.verify(false,true,"Time for Wednessday is Needed to Create Timesheet");
}
thu = ExcelUtils.getRowDatas("Thu",EnvParams.Opco)
if((thu==null)||(thu=="")){ 
ValidationUtils.verify(false,true,"Time for Thursday is Needed to Create Timesheet");
}
fri= ExcelUtils.getRowDatas("Fri",EnvParams.Opco)
if((fri==null)||(fri=="")){ 
ValidationUtils.verify(false,true,"Time for Friday is Needed to Create Timesheet");
}
sat = ExcelUtils.getRowDatas("Sat",EnvParams.Opco)
sun = ExcelUtils.getRowDatas("Sun",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}

function CreateTimeSheet(){ 
TextUtils.writeLog("Timesheet Creation Started"); 
Indicator.PushText("waiting for window to reponse");
aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);
  
}

//aqTestCase.Begin("Job Creation", "zfj://CH1-30");
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateTimesheet";
Language = "";

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
//Log.Message(EnvParams.Opco)
//Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

ExcelUtils.setExcelName(workBook, sheetName, true);
z=0;
invisible_Startindex = "";
invisible_Endindex = 6;
visible_Startindex = 0;
visible_Endindex = "";
jobNumber,weekno,workID,Descrip,mon,tue,wed,thu,fri,sat,sun,EmployeeNumber,EmployeeName,startdate,enddate = "";


STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create Timesheet started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME);
getDetails();
goToJobMenuItem();

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
  }
  
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Timesheet");
  
sheetName = "CreateTimesheet";
selectWeek();
aqUtils.Delay(5000, Indicator.Text);
//Delay(5000);
WorkspaceUtils.closeAllWorkspaces();

//aqTestCase.End();
}

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
Delay(3000);
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
Client_Managt.ClickItem("|Time & Expenses");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Time & Expenses");
}

} 

aqUtils.Delay(15000, Indicator.Text);
//Delay(10000); 
ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
TextUtils.writeLog("Entering into Time & Expenses from Time & Expenses Menu");
}


function selectWeek(){ 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
EmployeeNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3).getText();
var Visiblestatus = true;
while(Visiblestatus){ 
  EmployeeNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3).getText();
  EmployeeName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2).getText();
  if((EmployeeNumber!="")&&(EmployeeNumber!=null)){ 
    aqUtils.Delay(2000, Indicator.Text);
//    Delay(2000);
    Visiblestatus = false;
  }
}
  var previousMonth = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 2);
  var nextMonth = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 3);
  var week1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 1)
  var week2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 2)
  var week3 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 3)
  var week4 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 4)
  var week5 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 5)
  var week6 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 6)
  var status = true;
  while(status)// If Year is need to check code need to start from here
  if(week1.getText().OleValue.toString()==weekno){ 
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 1);
//    day.Click();
    selectDay(1);
    break;
  }else if(week2.getText().OleValue.toString()==weekno){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 8);
//    day.Click();
    selectDay(8);
    break;
  }else if(week3.getText().OleValue.toString()==weekno){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 15);
//    day.Click();
    selectDay(15);
    break;
  }else if(week4.getText().OleValue.toString()==weekno){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 22);
//    day.Click();
    selectDay(22);
    break;
  }else if(week5.getText().OleValue.toString()==weekno){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 29);
//    day.Click();
    selectDay(29);
    break;
  }else if(week6.getText().OleValue.toString()==weekno){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 36);
//    day.Click();
    selectDay(36);
    break;
  }else{
  Log.Message("week6 :"+week6.getText());
  Log.Message("weekno :"+weekno); 
  if(week6.getText()>weekno){ 
previousMonth.HoverMouse();
ReportUtils.logStep_Screenshot("");
  previousMonth.Click();
  aqUtils.Delay(5000, Indicator.Text);
//  Delay(5000);  
  }
  if(week6.getText()<weekno){ 
nextMonth.HoverMouse();
ReportUtils.logStep_Screenshot("");
  nextMonth.Click();
  aqUtils.Delay(5000, Indicator.Text);
//  Delay(5000);  
  }
  }
  
}

function selectDay(startday){
weekseparate = false;
var days = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday);  
   for(i=1;i<7;i++){ 
   var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i).getText();
   if(day=="1"){
    invisible_Startindex = i;
    visible_Endindex = i-1;
//    Log.Message("invisible_Startindex :"+invisible_Startindex);
//    Log.Message("visible_Endindex :"+visible_Endindex);
   weekseparate = true;
   break;
   
}
}

if(!weekseparate){ 
var Rejectline = false;
  for(i=0;i<7;i++){
//  for(i=visible_Startindex;i<=visible_Endindex;i++){  

   var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i);
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
  var monday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  var tuesday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 3);
  var wednesday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 4);
  var thursday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 5);
  var friday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 6);
  var saturday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 7);
  var sunday = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 8);
aqUtils.Delay(4000, Indicator.Text);
//Delay(4000);
var create_Timesheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4).getText();
if(create_Timesheet=="Create Time Sheet"){
Sys.HighlightObject(create_Timesheet);
var create_Timesheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
create_Timesheet.HoverMouse();
ReportUtils.logStep_Screenshot(""); 
create_Timesheet.Click();
}
TextUtils.writeLog("Entering Timesheet Details");
    aqUtils.Delay(2000, Indicator.Text);
//    Delay(2000);
//startdate =   
    var linegrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    var linecount = linegrid.getItemCount()
    var linecount_Status = true;
    for(var gg=3;gg<11;gg++){  
//    Log.Message(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim())
    if((linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()=="")||(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()==null))
    linecount_Status = false;    
    }
    if(!Rejectline){  
    var addline = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
addline.HoverMouse();
ReportUtils.logStep_Screenshot("");
    addline.Click();
    }
    else{
    var Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    Registration.Click();
    }
    aqUtils.Delay(4000, Indicator.Text);
//    Delay(4000);
    var keep = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McPlainCheckboxView", "",1,60000);
    keep.Keys("[Tab][Tab][Tab]");
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    var Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McValuePickerWidget", "", 2,60000);
    Job.Click();
    if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_1(Job,"Job",jobNumber,"Job Number")
    Job
    }
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    var work = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
    work.Click();
    if(work.getText()!=workID){
    WorkspaceUtils.SearchByValue(work,"Work Code",workID,"WorkCode")
    work
    }
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var description = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    if(((Descrip!="")||(Descrip!=null))&&(description.getText()!=Descrip)){
    description.setText(Descrip);
    ValidationUtils.verify(true,true,"Description is Entered");
    }
    var Time_for_day = switchcase(0,6)
//    Log.Message("Time_for_day :"+Time_for_day.length);
    for(var kk=0;kk<7;kk++){
//    var dashboard = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2+kk).getText();
//    if(dashboard!="0:00"){ 
aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var days = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//    days.Keys(getRowData1(Time_for_day[z]));
//Log.Message(Time_for_day[kk])
//if((switchDayValue(Time_for_day[kk])!="")||(switchDayValue(Time_for_day[kk])!=null)){
if(((switchDayValue(Time_for_day[kk])!="")||(switchDayValue(Time_for_day[kk])!=null))&&(days.getText()!=switchDayValue(Time_for_day[kk]))){
days.Keys(switchDayValue(Time_for_day[kk]));
ValidationUtils.verify(true,true,"Time is Entered for "+Time_for_day[kk]);
}
    z++;
//    Delay(1000);
//      }    
     
}
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2).toolTipText=="Save Time Sheet Line")
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).ChildCount>2)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3).toolTipText=="Save Time Sheet Line")
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
    
    if(Save.isEnabled_2){
    Save.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Save.Click();
TextUtils.writeLog("Timesheet is Saved");
    aqUtils.Delay(5000, Indicator.Text);
//    Delay(5000);
    ValidationUtils.verify(true,true,"Line of Timesheet is Saved");
    if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption =="Time & Expenses - Registrations"){ 
      var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Time & Expenses - Registrations").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Ok.Click();
      aqUtils.Delay(5000, Indicator.Text);
//      Delay(5000);
    }
  var balance = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 9);
//  if(balance.getText()!="0:00"){ 
  if(balance.getText().OleValue.toString().trim().indexOf("-")!=-1){
//  ReportUtils.logStep("INFO","Balance is not Zero or Timesheet is not balanced");
//  Log.Message("Balance is not Zero or Timesheet is not balanced");
  ValidationUtils.verify(false,true,"Balance is not Zero or Timesheet is not balanced");
  }else{ 
//  ReportUtils.logStep("INFO","Timesheet is balanced");
//  Log.Message("Timesheet is balanced");
  ValidationUtils.verify(true,true,"Timesheet is balanced");
  TextUtils.writeLog("Timesheet is balanced");
  }
    var Submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(Submit);
    Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Submit.Click();
    ValidationUtils.verify(true,true,"Timesheet is Submit");
    TextUtils.writeLog("Timesheet is Submit");
    
//    sheetName = "JobCreation";
//    var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)  
//   var name =  LogReport_name(ExlArray,comapany,Job_group);
//    var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";
////    Log.Message("Notepad :"+notepadPath)
////      var notepadPath = Project.Path+EnvParams.path+EnvParams.OpcoNumber+".txt";
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management",EmployeeNumber)
      ExcelUtils.WriteExcelSheet("Timesheet Employee Name",EnvParams.Opco,"Data Management",EmployeeName)
      ExcelUtils.WriteExcelSheet("Timesheet Week No",EnvParams.Opco,"Data Management",weekno)
      
      TextUtils.writeLog("Timesheet is submitted by Employee No:"+EmployeeNumber);
      TextUtils.writeLog("Timesheet is submitted by Employee Name:"+EmployeeName);
      TextUtils.writeLog("Timesheet is Created for Week No:"+weekno);
//    TextUtils.writeDetails(notepadPath,"Timesheet Employee No ",EmployeeNumber);
//    TextUtils.writeDetails(notepadPath,"Timesheet Week No ",weekno);

}
}
else{ 
var Rejectline = false;
  for(i=0;i<7;i++){
//  for(i=visible_Startindex;i<=visible_Endindex;i++){  

   var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i);
   if((day.background==8689648)||(day.background==16777215)||(day.background==3405809)){
   if(day.background==3405809){
   Rejectline = true;
   ReportUtils.logStep("INFO","Selected week is already Rejected, Altering Infomation in Registration panel")
   }
     }else{ 
       ValidationUtils.verify(false,true,"Value is already Submitted or Approved");
     }
     } 

    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+visible_Startindex);
   day.HoverMouse();
ReportUtils.logStep_Screenshot("");
    day.Click();
    TextUtils.writeLog("Week has been Selected");
    aqUtils.Delay(4000, Indicator.Text);
//    Delay(4000);
var create_Timesheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4).getText();
if(create_Timesheet=="Create Time Sheet"){
Sys.HighlightObject(create_Timesheet);
var create_Timesheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
   create_Timesheet.HoverMouse();
ReportUtils.logStep_Screenshot("");
create_Timesheet.Click();
}
TextUtils.writeLog("Entering Timesheet Details");
    aqUtils.Delay(2000, Indicator.Text);
//    Delay(2000);

    var linegrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    var linecount = linegrid.getItemCount();
var linecount_Status = true;
    for(var gg=3;gg<11;gg++){  
//    Log.Message(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim())
    if((linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()=="")||(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()==null))
    linecount_Status = false;    
    }
    if(!Rejectline){   
    var addline = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    addline.HoverMouse();
ReportUtils.logStep_Screenshot("");
    addline.Click();
    }
    else{
    var Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    Registration.Click();
    }
    aqUtils.Delay(2000, Indicator.Text);
//    Delay(2000);
    var keep = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "");
    keep.Keys("[Tab][Tab][Tab]");
    aqUtils.Delay(5000, Indicator.Text);
//    Delay(1000);
    var Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McValuePickerWidget", "", 2,60000);
    Job.Click();
    if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_1(Job,"Job",jobNumber,"Job Number");
    Job
    }
//    WorkspaceUtils.SearchByValues_Col_1(Job,"Job",jobNumber,"Job Number")
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    var work = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
    work.Click();
    if(work.getText()!=workID){
    WorkspaceUtils.SearchByValue(work,"Work Code",workID,"WorkCode");
    work
    }
//    WorkspaceUtils.SearchByValue(work,"Work Code",workID,"WorkCode")
aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var description = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    if(((Descrip!="")||(Descrip!=null))&&(description.getText()!=Descrip)){
    description.setText(Descrip);
    ValidationUtils.verify(true,true,"Description is Entered");
    }
    var Time_for_day = switchcase(visible_Startindex,visible_Endindex)
    var z=0;
    for(var k=0;k<7;k++){ 
      if((k>=visible_Startindex)&&(k<=visible_Endindex)){
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var days = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//    days.Keys(getRowData1(Time_for_day[z]));
//Log.Message(Time_for_day[z])
if(((switchDayValue(Time_for_day[z])!="")||(switchDayValue(Time_for_day[z])!=null))&&(days.getText()!=switchDayValue(Time_for_day[z]))){
days.Keys(switchDayValue(Time_for_day[z]));
ValidationUtils.verify(true,true,"Time is Entered for "+Time_for_day[z]);
}
    z++;
      }
      else{ 
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09); 
      }
    }
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2).toolTipText=="Save Time Sheet Line")
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).ChildCount>2)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3).toolTipText=="Save Time Sheet Line")
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
    
    if(Save.isEnabled_2){
Save.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Save.Click();
    aqUtils.Delay(5000, Indicator.Text);
//    Delay(5000);
    ValidationUtils.verify(true,true,"Line of Timesheet is Saved");
    TextUtils.writeLog("Timesheet is Saved");
    if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption =="Time & Expenses - Registrations"){ 
      var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Time & Expenses - Registrations").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Ok.Click();
      aqUtils.Delay(5000, Indicator.Text);
//      Delay(5000);
    }
  var balance = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 9);
//  if(balance.getText()!="0:00"){ 
  if(balance.getText().OleValue.toString().trim().indexOf("-")!=-1){
  ValidationUtils.verify(false,true,"Balance is not Zero or Timesheet is not balanced");
  }else{ 
  ValidationUtils.verify(true,true,"Timesheet is balanced");
  TextUtils.writeLog("Timesheet is balanced");
  }
    var Submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(Submit);
Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Submit.Click();
    ValidationUtils.verify(true,true,"Timesheet is Submit");
    TextUtils.writeLog("Timesheet is Submit");
    }
    aqUtils.Delay(5000, Indicator.Text);
//    Delay(5000);
    
//  for(i=invisible_Startindex;i<=invisible_Endindex;i++){  
//
//   var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday+i);
//   if((day.background==8689648)||(day.background==16777215)||(day.background==3405809)){
//   if(day.background==3405809){
//   Rejectline = true;
//   ReportUtils.logStep("INFO","Selected week is already Rejected, Altering Infomation in Registration panel")
//   }
//     }else{ 
//       ValidationUtils.verify(false,true,"Value is already Submitted or Approved");
//     }
//     } 
    
    var nextMonth = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 3);
nextMonth.HoverMouse();
ReportUtils.logStep_Screenshot("");
    nextMonth.Click();
    TextUtils.writeLog("Navigating to next month to create Time sheet for that week");
    aqUtils.Delay(4000, Indicator.Text);
//    Delay(4000);
    var days = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 1+invisible_Startindex);  
    days.HoverMouse();
ReportUtils.logStep_Screenshot("");
    days.Click();
    TextUtils.writeLog("Week has been Selected");
    aqUtils.Delay(2000, Indicator.Text);
//    Delay(2000);
var create_Timesheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4).getText();
if(create_Timesheet=="Create Time Sheet"){
Sys.HighlightObject(create_Timesheet);
var create_Timesheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
create_Timesheet.HoverMouse();
ReportUtils.logStep_Screenshot("");
create_Timesheet.Click();
}
TextUtils.writeLog("Entering Timesheet Details");
    aqUtils.Delay(4000, Indicator.Text);
//    Delay(4000);
    var linegrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    var linecount = linegrid.getItemCount()
var linecount_Status = true;
    for(var gg=3;gg<11;gg++){  
    if((linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()!="")||(linegrid.getItem(0).getText_2(gg).OleValue.toString().trim()!=null))
    linecount_Status = false;    
    }
    if(!Rejectline){   
    var addline = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    addline.HoverMouse();
ReportUtils.logStep_Screenshot("");
    addline.Click();
    }
    else{
    var Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    Registration.Click();
    }
    aqUtils.Delay(4000, Indicator.Text);
//    Delay(4000);
    var keep = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "");
    keep.Keys("[Tab][Tab][Tab]");
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    var Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McValuePickerWidget", "", 2,60000);
    Job.Click();
if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_1(Job,"Job",jobNumber,"Job Number");
    Job
    }
//    WorkspaceUtils.SearchByValues_Col_1(Job,"Job",jobNumber,"Job Number")
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    var work = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).WaitSWTObject("McValuePickerWidget", "", 2,60000);
    work.Click();
    if(work.getText()!=workID){
    WorkspaceUtils.SearchByValue(work,"Work Code",workID,"WorkCode");
    work
    }
//    WorkspaceUtils.SearchByValue(work,"Work Code",workID,"WorkCode")
aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var description = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    if(((Descrip!="")||(Descrip!=null))&&(description.getText()!=Descrip)){
    description.setText(Descrip);
    ValidationUtils.verify(true,true,"Description is Entered");
    }
    var Time_for_day = switchcase(invisible_Startindex,invisible_Endindex)
    var z=0;
    for(var k=0;k<7;k++){ 
      if((k>=invisible_Startindex)&&(k<=invisible_Endindex)){ 
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    var days = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//    days.Keys(getRowData1(Time_for_day[z]));
if(((switchDayValue(Time_for_day[z])!="")||(switchDayValue(Time_for_day[z])!=null))&&(days.getText()!=switchDayValue(Time_for_day[z]))){
days.Keys(switchDayValue(Time_for_day[z]));
ValidationUtils.verify(true,true,"Time is Entered for "+Time_for_day[z]);
}
    z++;
      }
      else{
    aqUtils.Delay(1000, Indicator.Text);
//    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09); 
      }
    }
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2).toolTipText=="Save Time Sheet Line")
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).ChildCount>2)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3).toolTipText=="Save Time Sheet Line")
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
    
    if(Save.isEnabled_2){
Save.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Save.Click();
    aqUtils.Delay(5000, Indicator.Text);
//    Delay(5000);
    ValidationUtils.verify(true,true,"Line of Timesheet is Saved");
    TextUtils.writeLog("Timesheet is Saved");
    
    if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption =="Time & Expenses - Registrations"){ 
      var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Time & Expenses - Registrations").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Ok.Click();
      aqUtils.Delay(5000, Indicator.Text);
//      Delay(5000);
    }
    
    var Submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(Submit);
Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Submit.Click();
    ValidationUtils.verify(true,true,"Timesheet is Submit");
    TextUtils.writeLog("Timesheet is Submit");
   }
  var balance = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 9);
//  if(balance.getText()!="0:00"){
  if(balance.getText().OleValue.toString().trim().indexOf("-")!=-1){
  ValidationUtils.verify(false,true,"Balance is not Zero or Timesheet is not balanced");
  }else{ 
  ValidationUtils.verify(true,true,"Timesheet is balanced");
  TextUtils.writeLog("Timesheet is balanced");
  }
  
//  sheetName = "JobCreation";
//  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)  
//   var name =  LogReport_name(ExlArray,comapany,Job_group);
//    var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";
////    Log.Message("Notepad :"+notepadPath)
////      var notepadPath = Project.Path+EnvParams.path+EnvParams.OpcoNumber+".txt";
//    TextUtils.writeDetails(notepadPath,"Timesheet Employee No ",EmployeeNumber);
//    TextUtils.writeDetails(notepadPath,"Timesheet Week No ",weekno);
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management",EmployeeNumber)
      ExcelUtils.WriteExcelSheet("Timesheet Employee Name",EnvParams.Opco,"Data Management",EmployeeName)
      ExcelUtils.WriteExcelSheet("Timesheet Week No",EnvParams.Opco,"Data Management",weekno)
      
      TextUtils.writeLog("Timesheet is submitted by Employee No:"+EmployeeNumber);
      TextUtils.writeLog("Timesheet is submitted by Employee Name:"+EmployeeName);
      TextUtils.writeLog("Timesheet is Created for Week No:"+weekno);
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



