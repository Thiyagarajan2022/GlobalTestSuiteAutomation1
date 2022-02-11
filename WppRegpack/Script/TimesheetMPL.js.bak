//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Data Management";
var Language = "";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");
ExcelUtils.setExcelName(workBook, sheetName, true);
var empNumber, empName, weekNo ,TYear, timeSheetHours,globalClientName,jobNumber = "";
var pdflineSplit ="";

function MPLTimesheet(){
  TextUtils.writeLog("MPL Timesheet Started"); 
Indicator.PushText("waiting for window to open");

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}


excelName = EnvParams.path;
workBook = Project.Path+excelName;


ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "MPL Timesheet started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 

getDetails();
goToTimeAndExpenseMenuItem();  
selectWeek();
print();
validateTimesheetPDF();
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
 
  ExcelUtils.setExcelName(workBook, sheetName, true);
  empName = ReadExcelSheet("Timesheet Employee Name",EnvParams.Opco,"Data Management");  
  if((empName=="")||(empName==null))
  ValidationUtils.verify(false,true,"Employee Name is needed to validate Timesheet MPL");
    
  empNumber = ReadExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management");  
  if((empNumber=="")||(empNumber==null))
  ValidationUtils.verify(false,true,"Timesheet Employee No is needed to validate Timesheet MPL"); 
 
  weekNo = ReadExcelSheet("Timesheet Week No",EnvParams.Opco,"Data Management");
  if((weekNo=="")||(weekNo==null))
  ValidationUtils.verify(false,true,"Week No is needed to validate Timesheet MPL");
  
  jobNumber = ReadExcelSheet("Timesheet Job No",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"JobNumber is needed to validate Timesheet MPL")
  
  timeSheetHours = ReadExcelSheet("Timesheet Hours",EnvParams.Opco,"Data Management");
  if((timeSheetHours=="")||(timeSheetHours==null))
  ValidationUtils.verify(false,true,"TimeSheetHours is needed to validate Timesheet MPL");
  
  globalClientName = ReadExcelSheet("Global Client Name",EnvParams.Opco,"Data Management");
  if((globalClientName=="")||(globalClientName==null))
  ValidationUtils.verify(false,true,"Global Client Name is needed to validate Timesheet MPL");
  
  ExcelUtils.setExcelName(workBook, "ApproveTimesheet", true);
  TYear = ExcelUtils.getRowDatas("Year",EnvParams.Opco)
  if((TYear=="")||(TYear==null))
  ValidationUtils.verify(false,true,"Year is needed to validate Timesheet MPL");

}



function goToTimeAndExpenseMenuItem(){

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.TimeExpense.Exists()){
 ImageRepository.ImageSet.TimeExpense.Click();
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
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
}

} 

aqUtils.Delay(10000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
TextUtils.writeLog("Entering into Time & Expenses from Time & Expenses Menu");
}


function selectWeek(){ 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var previousMonth = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 2);
  var nextMonth = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 3);
  var week1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 1)
  var week2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 2)
  var week3 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 3)
  var week4 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 4)
  var week5 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 5)
  var week6 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 3).SWTObject("DateChooser$CellLabel", "", 6)
  var YearMonth = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("Label", "*");
  var previousYear = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 1);
  var nextYear = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 1).SWTObject("McGenericButton", "", 4);
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
  if(week1.getText().OleValue.toString()==weekNo){ 
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 1);
//    day.Click();
    selectDay(1);
    break;
  }else if(week2.getText().OleValue.toString()==weekNo){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 8);
//    day.Click();
    selectDay(8);
    break;
  }else if(week3.getText().OleValue.toString()==weekNo){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 15);
//    day.Click();
    selectDay(15);
    break;
  }else if(week4.getText().OleValue.toString()==weekNo){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 22);
//    day.Click();
    selectDay(22);
    break;
  }else if(week5.getText().OleValue.toString()==weekNo){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 29);
//    day.Click();
    selectDay(29);
    break;
  }else if(week6.getText().OleValue.toString()==weekNo){
    status = false;
    var day = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", 36);
//    day.Click();
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
  
}

function print(){
  var print = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl2;
  Sys.HighlightObject(print);
    waitForObj(print)
    ReportUtils.logStep_Screenshot();
    print.Click();
TextUtils.writeLog("Print Timesheet is Clicked and saved"); 
aqUtils.Delay(5000, Indicator.Text);        
    
var SaveTitle = "";
var sFolder = "";
var pdf = 
Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Time Sheet"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
  if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Time Sheet"+"*", 1).WndCaption.indexOf("Print Time Sheet")!=-1){
    aqUtils.Delay(5000, Indicator.Text);
WorkspaceUtils.waitForObj(pdf);
Sys.HighlightObject(pdf)
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
    
if(ImageRepository.PDF.ChooseFolder.Exists())
ImageRepository.PDF.ChooseFolder.Click();
else{ 
var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
WorkspaceUtils.waitForObj(window);

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x73); //F4
Sys.Desktop.KeyUp(0x12); //Alt
Sys.Desktop.KeyUp(0x73); //F4
aqUtils.Delay(2000, Indicator.Text);
Sys.HighlightObject(pdf)

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
}
var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
aqUtils.Delay(2000, Indicator.Text);
SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");

var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
  saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print TimeSheet is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("TimeSheetMPL",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")
    
}

function validateTimesheetPDF()
{
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  var fileName = ExcelUtils.getRowDatas("TimeSheetMPL",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"TimeSheetMPL is needed to validate");
  }

  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }

 
  pdflineSplit = docObj.split("\r\n");
       
  if(EnvParams.Country.toUpperCase()=='CHINA')
    Language = "Chinese (Simplified)";
    
   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "TIME SHEET").OleValue.toString().trim());
  
       if(index>=0){
          ReportUtils.logStep("INFO","Heading is available in Timesheet PDF")
          ValidationUtils.verify(true,true,"Heading is available in Timesheet PDF")
          TextUtils.writeLog("Heading is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Heading is not available in Timesheet PDF")
          
   var index = docObj.indexOf(timeSheetHours+".00");
    if(index>=0){
          ReportUtils.logStep("INFO",timeSheetHours+" TotalHours is matching with Pdf");
          ValidationUtils.verify(true,true,timeSheetHours+" TotalHours is matching with Pdf")
          TextUtils.writeLog(timeSheetHours+" TotalHours is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"TotalHours is not same in Timesheet PDF");      
  
  
  verifyEmployeeNameAndNumber();
  verifyWeekNo();     
  verifyJobNumber();    
  verifyGlobalClientName(); 
  
}


function verifyEmployeeNameAndNumber()
{
    var notFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim()))
        {
         if(pdflineSplit[j].includes(empName+"  -  "+empNumber))
             {
             Log.Message(empName+"  -  "+empNumber+" empName and Number is matching with Pdf"+empName+"  -  "+empNumber);             
             ValidationUtils.verify(true,true,"empName and Number is matching with Pdf:");
             TextUtils.writeLog("empName and Number is matching with Pdf:"+empName+"  -  "+empNumber);
             notFound = true;
             break;
             }
             }
         if(j==pdflineSplit.length-1 && !notFound)
          ValidationUtils.verify(false,true,"empName and Number is not same/found in TimesheetPdf");
        }  
}

function verifyWeekNo()
{
    var weekNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Week No").OleValue.toString().trim()))
        {
         if(pdflineSplit[j].includes(weekNo))
             {
             Log.Message(weekNo+" weekNo is matching with Pdf");             
             ValidationUtils.verify(true,true,"weekNo is matching with Pdf:"+weekNo);
             TextUtils.writeLog("weekNo is matching with Pdf:"+weekNo);
             weekNoFound = true;
             break;
             }
             }
         if(j==pdflineSplit.length-1 && !weekNoFound)
          ValidationUtils.verify(false,true,"weekNo is not same/found in TimesheetPdf");
        }  
}
 
function verifyJobNumber()
{
    var jobNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(jobNumber))
             {
             Log.Message(jobNumber+" jobNumber is matching with Pdf");
             ValidationUtils.verify(true,true,"jobNumber is matching with Pdf:"+jobNumber);
             TextUtils.writeLog("jobNumber is matching with Pdf:"+jobNumber);
             jobNoFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !jobNoFound)
          ValidationUtils.verify(false,true,"jobNumber is not same/found in TimesheetPdf");
        }  
}
 
function verifyGlobalClientName()
{
    var globalClientNameFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(globalClientName.includes(pdflineSplit[j]))
             {
             Log.Message(globalClientName+" globalClientName is matching with Pdf");
             ValidationUtils.verify(true,true,"global ClientName is matching with Pdf:"+globalClientName);
             TextUtils.writeLog("global ClientName is matching with Pdf:"+globalClientName);
             globalClientNameFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !globalClientNameFound)
          ValidationUtils.verify(false,true,"globalClientName is not same/found in TimesheetFile");
        }  
        
}



function selectDay(startday){
var days = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McCalendarWidget", "", 2).SWTObject("McDateChooser", "").SWTObject("McComposite", "", 2).SWTObject("Composite", "", 4).SWTObject("DateChooser$CellLabel", "", startday);

days.HoverMouse();
ReportUtils.logStep_Screenshot("");  
days.Click();  
TextUtils.writeLog("Week has been Selected to get Approver");
}
 