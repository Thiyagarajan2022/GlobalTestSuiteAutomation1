﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart
//var excelName = EnvParams.getEnvironment();

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "QueryExistingEmployee";
var FullName,EmployeeNo,Gender,Country,DateEmployed,Position,Email,ApproverGroup,EmploymentType,SalesEmployee,WeekCalendarNo = "";
var Language = "";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var empNumber = "";

var approvers;
var STIME="";

function goToMenu(){ 
        ReportUtils.logStep_Screenshot("Go To menu");
 TextUtils.writeLog("QueryExistingEmployee started"); 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
  menuBar.DblClick();

if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();  
}


//if(ImageRepository.ImageSet.Emp_1.Exists()){
//ImageRepository.ImageSet.Emp_1.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.Emp_2.Exists()){
//ImageRepository.ImageSet.Emp_2.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.Emp_3.Exists()){
//ImageRepository.ImageSet.Emp_3.DblClick();// GL
//}

  aqUtils.Delay(3000, Indicator.Text);
//  Delay(3000);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
//  Delay(2000);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
  
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
Client_Managt.ClickItem("|Employees");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Employees");
}
}
        ReportUtils.logStep_Screenshot("Navigate to Employess");
 TextUtils.writeLog("Navigate Employees"); 

Delay(3000);
var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1)
.SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Employees");
All_Emp.Click();
//EmployeeNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3).getText();
//var Visiblestatus = true;

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(EnvParams.Opco);
//  firstCell.Keys("1707");
firstCell.Keys("[Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
Log.Message(EmployeeNo);
//  job.setText("GAIL C COUTINHO")
job.setText(EmployeeNo);

Delay(6000);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(EmployeeNo)){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
  
    
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Employee Created is available in system");
   
  
if(flag){ 
closefilter.Click();
Delay(10000);

}


var employeename = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.EmpNamevalue;

var employeeNo = NameMapping.Sys_new.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.Composite.McTextWidget;
var employeeCountry = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.EmpCountry
var employeeEmail=Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.EmpEmail
var employeePosition = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.EmpPosition
var employeecompany = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.EmpCompany
var empDepartment =Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.EmpDept
var employeeBusinessUnit =Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.BusinessUnit;
var employeeRole = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.EmployeeRole;
var employeeSupervisor = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.EmpSupervisor;

Sys.HighlightObject(employeename);
Sys.HighlightObject(employeeNo);
Sys.HighlightObject(employeeCountry);
Sys.HighlightObject(employeeEmail);
Sys.HighlightObject(employeePosition);
Sys.HighlightObject(employeecompany);
Sys.HighlightObject(empDepartment);
Sys.HighlightObject(employeeBusinessUnit);
Sys.HighlightObject(employeeRole);
Sys.HighlightObject(employeeSupervisor);

      ReportUtils.logStep_Screenshot("Validating Employee Details");
 TextUtils.writeLog("Validating Employee Details"); 

      if(employeename.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "Employeename is present");
            ValidationUtils.verify(true,true,"Employeename is present");
            Log.Checkpoint("Employeename is present")
        }
        else{
           ValidationUtils.verify(false,true,"Employeename is not present");
        }
        
         if(employeeNo.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "Employee No is present");
            ValidationUtils.verify(true,true,"Employee No is present");
            Log.Checkpoint("Employee No is present")
        }
        else{
           ValidationUtils.verify(false,true,"Employee No is not present");
        }
        
         if(employeeCountry.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "EmployeeCountry is present");
            ValidationUtils.verify(true,true,"EmployeeCountry is present");
            Log.Checkpoint("EmployeeCountry is present")
        }
        else{
           ValidationUtils.verify(false,true,"employeename is not present");
        }
        
         if(employeeEmail.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "EmployeeEmail is present");
            ValidationUtils.verify(true,true,"EmployeeEmail Matched Correctly");
            Log.Checkpoint("EmployeeEmail is present")
        }
        else{
           ValidationUtils.verify(false,true,"EmployeeEmail is present");
        }
         if(employeePosition.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "employeePosition is present");
            ValidationUtils.verify(true,true,"employeePosition is present");
            Log.Checkpoint("employeePosition is present")
        }
        else{
           ValidationUtils.verify(false,true,"employeePosition is present");
        }
         if(employeecompany.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "employeecompany is present");
            ValidationUtils.verify(true,true,"employeecompany is present");
            Log.Checkpoint("employeecompany is present")
        }
        else{
           ValidationUtils.verify(false,true,"employeecompany is present");
        }
         if(empDepartment.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "empDepartment is present");
            ValidationUtils.verify(true,true,"empDepartment is present");
            Log.Checkpoint("empDepartment is present")
        }
        else{
           ValidationUtils.verify(false,true,"empDepartment is present");
        }
         if(employeeBusinessUnit.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "employeeBusinessUnit is present");
            ValidationUtils.verify(true,true,"employeeBusinessUnit is present");
              Log.Checkpoint("employeeBusinessUnit is present")
        }
        else{
           ValidationUtils.verify(false,true,"employeeBusinessUnit is present");
        }
         if(employeeRole.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "employeeRole is present");
            ValidationUtils.verify(true,true,"employeeRole is present");
              Log.Checkpoint("employeeRole is present")
        }
        else{
           ValidationUtils.verify(false,true,"employeeRole is present");
        }
         if(employeeSupervisor.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "employeeSupervisor is present");
            ValidationUtils.verify(true,true,"employeeSupervisor is present");
              Log.Checkpoint("employeeSupervisor is present")
        }
        else{
           ValidationUtils.verify(false,true,"employeeSupervisor is present");
        }
}

function test()
{
  var employeename = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.EmpNamevalue;
    if(employeename.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "employeename is present");
            ValidationUtils.verify(true,true,"Product Matched Correctly");
        }
        else{
           ValidationUtils.verify(false,true,"employeename is present");
        }
}

function QueryExistingEmployee(){
level=0;
EmployeeNo="";
  //var workBook = "C:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\DS_CHN_SYSTEST.xlsx"
var sheetName = "QueryExistingEmployee";
ExcelUtils.setExcelName(workBook, sheetName, true);
  EmployeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
  if((EmployeeNo=="")||(EmployeeNo==null)){
//  jobNumber = readlog();
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  EmployeeNo = ExcelUtils.ReadExcelSheet("Employee No",EnvParams.Opco,"Data Management");
  }
  getDetails();
  goToMenu();
  //employeeinfo();
  
WorkspaceUtils.closeAllWorkspaces();


  }
  
//  if((EmployeeNo=="")||(EmployeeNo==null))
//  ValidationUtils.verify(false,true,"Job Number is needed to Change the Employee");
//  Log.Message(EmployeeNo)
//
//  ExcelUtils.setExcelName(workBook, sheetName, true);
//  EmployeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
//  if((EmployeeNo=="")||(EmployeeNo==null)){
////  jobNumber = readlog();
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  EmployeeNo = ReadExcelSheet("Employee No",EnvParams.Opco,"Data Management");
//  }
//  
//  if((EmployeeNo=="")||(EmployeeNo==null))
//  ValidationUtils.verify(false,true,"Job Number is needed to Change the Employee");
//
//}
function goToEmployees(){ 

//  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
//  employees.DblClickItem("|Employees");

//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//var Add_Visible0 = true;
//var New_Employee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 3,60000);
//while(Add_Visible0){
//if(New_Employee.isEnabled()){
//New_Employee.HoverMouse();
//ReportUtils.logStep_Screenshot();
//New_Employee.Click();
//Add_Visible0 = false;
//}
//}
  }
  
  //getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
//FullName = ExcelUtils.getRowDatas("FullName",EnvParams.Opco)
//if((FullName==null)||(FullName=="")){ 
//ValidationUtils.verify(false,true,"FullName is Needed to Create a Employee");
//}
//Gender = ExcelUtils.getRowDatas("Gender",EnvParams.Opco)
//if((Gender==null)||(Gender=="")){ 
//ValidationUtils.verify(false,true,"Gender is Needed to Create a Employee");
//}
//Country = ExcelUtils.getRowDatas("Country",EnvParams.Opco)
//if((Country==null)||(Country=="")){ 
//ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
//}

//comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
//if((comapany==null)||(comapany=="")){ 
//ValidationUtils.verify(false,true,"Company Number is Needed to Create a Employee");
//}
//DateEmployed = ExcelUtils.getRowDatas("DateEmployed",EnvParams.Opco)
//if((DateEmployed==null)||(DateEmployed=="")){ 
//ValidationUtils.verify(false,true,"DateEmployed is Needed to Create a Employee");
//}
//Position = ExcelUtils.getRowDatas("Position",EnvParams.Opco)
//if((Position==null)||(Position=="")){ 
//ValidationUtils.verify(false,true,"Position is Needed to Create a Employee");
//}
//Email = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
//if((Email==null)||(Email=="")){ 
//ValidationUtils.verify(false,true,"Email is Needed to Create a Employee");
//}
//ApproverGroup = ExcelUtils.getRowDatas("ApproverGroup",EnvParams.Opco)
//
//EmploymentType= ExcelUtils.getRowDatas("EmploymentType",EnvParams.Opco)
//if((EmploymentType==null)||(EmploymentType=="")){ 
//ValidationUtils.verify(false,true,"EmploymentType is Needed to Create a Employee");
//}
//SalesEmployee= ExcelUtils.getRowDatas("Sales Employee",EnvParams.Opco)
//if((SalesEmployee==null)||(SalesEmployee=="")){ 
//ValidationUtils.verify(false,true,"Sales Employee is Needed to Create a Employee");
//}
//EmployeeDepartment_Name= ExcelUtils.getRowDatas("EmployeeDepartment_Name",EnvParams.Opco)
//if((EmployeeDepartment_Name==null)||(EmployeeDepartment_Name=="")){ 
//ValidationUtils.verify(false,true,"EmployeeDepartment_Name is Needed to Create a Employee");
//}
//EmployeeCostCentre_Name= ExcelUtils.getRowDatas("EmployeeCostCentre_Name",EnvParams.Opco)
//if((EmployeeCostCentre_Name==null)||(EmployeeCostCentre_Name=="")){ 
//ValidationUtils.verify(false,true,"EmployeeCostCentre_Name is Needed to Create a Employee");
//}
//Supervisor= ExcelUtils.getRowDatas("Supervisor",EnvParams.Opco)
//if((Supervisor==null)||(Supervisor=="")){ 
//ValidationUtils.verify(false,true,"Supervisor is Needed to Create a Employee");
//}
//AbsenceApprover= ExcelUtils.getRowDatas("AbsenceApprover",EnvParams.Opco)
//if((AbsenceApprover==null)||(AbsenceApprover=="")){ 
//ValidationUtils.verify(false,true,"AbsenceApprover is Needed to Create a Employee");
//}
//Role= ExcelUtils.getRowDatas("Role",EnvParams.Opco)
//if((Role==null)||(Role=="")){ 
//ValidationUtils.verify(false,true,"Role is Needed to Create a Employee");
//}
//VacationCalendar= ExcelUtils.getRowDatas("Vacation Calendar",EnvParams.Opco)
//if((VacationCalendar==null)||(VacationCalendar=="")){ 
//ValidationUtils.verify(false,true,"Vacation Calendar is Needed to Create a Employee");
//}
//WeekCalendarNo= ExcelUtils.getRowDatas("Week Calendar No.",EnvParams.Opco)
//if((WeekCalendarNo==null)||(WeekCalendarNo=="")){ 
//ValidationUtils.verify(false,true,"Week Calendar No. is Needed to Create a Employee");
//}

}





function employeeinfo(){
  
//var FullName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 2)
//Sys.HighlightObject(FullName);
//var Gender = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2)
//Sys.HighlightObject(Gender);
//var Country = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McPopupPickerWidget", "", 2)
//Sys.HighlightObject(Country);
  
  if(FullName!=""){
    Log.Message(FullName);
  var employee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1203 Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 2)
 employee.Click();
  Sys.Desktop.KeyDown(0x08)
  Sys.Desktop.KeyUp(0x08)
  employee.setText(FullName + " "+STIME );  
  Delay(1000);
Log.Message(FullName);  
ValidationUtils.verify(true,true,"Employee Name is updated in Maconomy");
ReportUtils.logStep_Screenshot();
}

//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
if(Gender!=""){
var Gender1 = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2)
Sys.HighlightObject(Gender1);
Gender1.Click();
  Gender1.Keys(Gender);  
Delay(1000);
Log.Message(Gender)  
ValidationUtils.verify(true,true,"Gender is updated in Maconomy"); 
ReportUtils.logStep_Screenshot();
}

if(Country!=""){
  var Country1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2)
Sys.HighlightObject(Country1);
Country1.Click();
//Sys.Process("Maconomy").Refresh();
Country1.Keys(Country);  
Delay(1000);
Log.Message(Country);
ValidationUtils.verify(true,true,"Country is updated in Maconomy");
ReportUtils.logStep_Screenshot();
}


if(Email!=""){
var Email_1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(Email_1);
Email_1.Click();
  Sys.Desktop.KeyDown(0x08)
  Sys.Desktop.KeyUp(0x08)
  Email_1.setText(Email + " "+STIME );  
  Delay(1000);
Log.Message(Email);  
ValidationUtils.verify(true,true,"Email is updated in Maconomy");
ReportUtils.logStep_Screenshot();
}


if(Position!=""){  
  //Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
 var Position_1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);Sys.HighlightObject(Position_1);
Position_1.Click();
Sys.Desktop.KeyDown(0x08)
  Sys.Desktop.KeyUp(0x08)
Position_1.setText(Position);
Delay(1000);
Log.Message(Position); 
ValidationUtils.verify(true,true,"Position is updated in Maconomy"); 
ReportUtils.logStep_Screenshot();
}
//  
//  if(VacationCalendar!=""){
//var VacationCalendar_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McValuePickerWidget", "", 2);
//VacationCalendar_1.Click();
// WorkspaceUtils.SearchByValue(VacationCalendar_1,"Vacation Calendar",VacationCalendar,"Vacation Calendar");
//     }else{ 
//  ValidationUtils.verify(false,true,"Vacation Calendar is Needed to Create a Employee");
//}
//
//    
//var WeekCalendarNo_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McValuePickerWidget", "", 2);
//if(WeekCalendarNo!=""){
// WeekCalendarNo_1.Click();
//    WorkspaceUtils.SearchByValue(WeekCalendarNo_1,"Week Calendar",WeekCalendarNo,"WeekCalendar No");
//
//   }else{ 
//  ValidationUtils.verify(false,true,"Week Calendar No is Needed to Create a Employee");
//}

 var save  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  save.Click();
  Delay(5000);
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
  
  
//var Refresh1= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
 var submit =Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
Sys.HighlightObject(submit);
submit.HoverMouse();
ReportUtils.logStep_Screenshot();
submit.Click();
Delay(6000); 
 ReportUtils.logStep("INFO", "Changes has Submitted");
 ValidationUtils.verify(true,true,"Employee Details has Changed");
 
//  var submit =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
//  submit.Click();
//  Delay(5000);
//  ReportUtils.logStep("INFO", "Changes has Submitted");
//  ValidationUtils.verify(true,true,"Employee Details has Changed");
   
  
var approve_Bar = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
//Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
Sys.HighlightObject(approve_Bar);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
Delay(2000);
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
var All_approver = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
Delay(1000);
All_approver.Click();
Delay(3000);
ReportUtils.logStep_Screenshot();
var approver_table = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
Log.Message(approver_table.getItemCount());
var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   var approvers="";
   if(approver_table.getItem(z).getText_2(8)!="Approved"){
   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   ReportUtils.logStep("INFO","Employee Approver level : " +z+ " Approver :" +approvers);
//   Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;

   Approve_Level[y] = EnvParams.Opco+"*"+EmployeeNo+"*"+approvers;
   Log.Message(Approve_Level[y])
   y++;
   }
}
}
}

var info_Bar = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
Sys.HighlightObject(info_Bar)
info_Bar.Click();
Delay(4000);

ImageRepository.ImageSet.Forward.Click();
Delay(4000);
var OpCo1 = EnvParams.Opco;
Log.Message(OpCo1);
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
Log.Message(OpCo2);
if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
level = 1;
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
//Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 8)
//Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 8)
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
  
  }
  }
  }
  

function Employee_Information(){ 
Delay(5000);
var Name = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("McTextWidget", "", 2,60000);
var Add_Visible = true;
while(Add_Visible){
if(Name.isEnabled()){
Delay(2000);
Add_Visible = false;
  
if(FullName!=""){
var FullName_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
FullName_1.Click();
FullName_1.setText(FullName + " "+STIME );
ValidationUtils.verify(true,true,"Employee Name is entered in Maconomy");
}else{ 
  ValidationUtils.verify(false,true,"Employee Name is Needed to Create a Employee");
}
  
if(Gender!=""){
var Gender_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
Gender_1.Click();
Gender_1.Keys(Gender);
Delay(5000);
ValidationUtils.verify(true,true,"Gender is selected in Maconomy"); 

}

var Country_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2)
if(Country!=""){
Country_1.Click();

Sys.Process("Maconomy").Refresh();
WorkspaceUtils.DropDownList(Country,"Country");
}else{ 
  ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
}
  
  

 
var Company_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(comapany!=""){
Company_1.Click();
var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
//Log.Message(Language)
WorkspaceUtils.config_with_Maconomy_Validation(Company_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),comapany,ExlArray,"Company Number");
}
else{ 
  ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
}



if(DateEmployed!=""){
var DateEmployed_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2)
WorkspaceUtils.CalenderDateSelection(DateEmployed_1,DateEmployed)
ValidationUtils.verify(true,true,"Date Employed is selected in Maconomy"); 
}else{ 
  ValidationUtils.verify(false,true,"Date Employed is Needed to Create a Employee");
}

//if(Employee_detail[5]!=""){
//var TerminateDate = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2)
//WorkspaceUtils.CalenderDateSelection(TerminateDate,Employee_detail[5])
////  DateEmployed.setText(Employee_detail[5]);
//}
  
if(Position!=""){  
var Position_1 =  Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
Position_1.setText(Position);
ValidationUtils.verify(true,true,"Position is Entered in Maconomy"); 
}else{ 
  ValidationUtils.verify(false,true,"Position is Needed to Create a Employee");
}
  
if(Email!=""){ 
var Email_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
var Eml_split1 = Email.substring(0,Email.indexOf("@"));
var Eml_split2 = Email.substring(0,Email.indexOf("@"));
Eml_split1 = Eml_split1 +" "+STIME;
Eml_split1 = Eml_split1.replace(/[_: ]/g,"");
Email = Eml_split1+Eml_split2
Email_1.setText(Email);  
ValidationUtils.verify(true,true,"Email Id is Entered in Maconomy"); 
}else{ 
  ValidationUtils.verify(false,true,"Email Id is Needed to Create a Employee");
}

  
var ApproverGroup_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2) 
if(ApproverGroup!=null){
ApproverGroup_1.Click();

Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible3 = true;
 while(Add_Visible3){
if(list.isEnabled()){
Add_Visible3 = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==ApproverGroup){ 
          Delay(1000);

            
          list.Keys("[Enter]");
          Delay(3000);
          ValidationUtils.verify(true,true,"Approver Group is selected in Maconomy");
          break;
        }else{ 
        if(i!=0){
         Log.Message(list.getItem(i).getText_2(0));
        list.Keys("[Down]");
          }
        }
          
      }else{ 
      if(i!=0){
         Log.Message(list.getItem(i).getText_2(0));
        list.Keys("[Down]");
          }
        
      }
    }
}
}
  
}
  
var EmploymentType_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2);
if(EmploymentType!=""){
EmploymentType_1.Click();  
Delay(5000);
Sys.Process("Maconomy").Refresh(); 
WorkspaceUtils.DropDownList(EmploymentType,"Employment Type"); 

 }else{ 
  ValidationUtils.verify(false,true,"Employment Type is Needed to Create a Employee");
} 
  
  
var SalesEmployee_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
if((SalesEmployee!="")&&(SalesEmployee!=null)){
if(SalesEmployee.toUpperCase()=="YES"){ 
if(!SalesEmployee_1.getSelection()){ 
  SalesEmployee_1.Click();
    ValidationUtils.verify(true,true,"Create Sales Employee checkBox is Clicked");
    }
  }else{ 
if(SalesEmployee_1.getSelection()){ 
  SalesEmployee_1.Click();
    ValidationUtils.verify(false,true,"Create Sales Employee checkBox is UnClicked");
    }
  }
  }
else{ 
if(SalesEmployee_1.getSelection()){ 
  SalesEmployee_1.Click();
    ValidationUtils.verify(false,true,"Create Sales Employee checkBox is UnClicked");
    }
  }


var EmployeeDepartment_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(EmployeeDepartment_Name!=""){
EmployeeDepartment_1.Click();
var ExlArray = getExcelData("ValidateDepartment",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation_Name_column2(EmployeeDepartment_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 3").OleValue.toString().trim(),EmployeeDepartment_Name,ExlArray,"Employee Department Number");
 }else{ 
  ValidationUtils.verify(false,true,"Employment Type is Needed to Create a Employee");
} 

    
var EmployeeCostCentre_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(EmployeeCostCentre_Name!=""){
EmployeeCostCentre_1.Click();
var ExlArray = getExcelData("Validate_BusinessUnit",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation_Name_column2(EmployeeCostCentre_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 5").OleValue.toString().trim(),EmployeeCostCentre_Name,ExlArray,"Business Unit Number");
}else{ 
  ValidationUtils.verify(false,true,"Employee Cost Centre is Needed to Create a Employee");
}

    
var Supervisor_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Supervisor!=""){
Supervisor_1.Click();
WorkspaceUtils.SearchByValue(Supervisor_1,"Employee",Supervisor,"Supervisor");

     }else{ 
  ValidationUtils.verify(false,true,"Supervisor Centre is Needed to Create a Employee");
} 

    
    
var AbsenceApprover_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(AbsenceApprover!=""){
AbsenceApprover_1.Click();
WorkspaceUtils.SearchByValue(AbsenceApprover_1,"Employee",AbsenceApprover,"AbsenceApprover");

     }else{ 
  ValidationUtils.verify(false,true,"Absence Approver is Needed to Create a Employee");
}
    
    
//var Secretary =  Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
//if(Employee_detail[14]!=""){
//Secretary.Click();
//WorkspaceUtils.SearchByValue(Secretary,"Employee",Employee_detail[14]);
//
//}

  Delay(2000);
  var next = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >")
 Sys.HighlightObject(next);
 next.HoverMouse();
ReportUtils.logStep_Screenshot();
 next.Click();


}
}
}


function Employee_Information1(){ 
Delay(2000);
var Role_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
if(Role!=""){
 Role_1.Click();
var ExlArray = getExcelData("Validate_EmployeeCategories",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(Role_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee Category").OleValue.toString().trim(),Role,ExlArray,"Role");
// WorkspaceUtils.SearchByValue(Role_1,"Employee Category",Role,"Role");
     }else{ 
  ValidationUtils.verify(false,true,"Role is Needed to Create a Employee");
} 

    
//if(Employee_detail[16]!=""){
//var MustUseTimeSheets = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
//MustUseTimeSheets.Click();
//
//Sys.Process("Maconomy").Refresh();
//WorkspaceUtils.DropDownList(Employee_detail[16]);
//}
  
var VacationCalendar_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McValuePickerWidget", "", 2);
if(VacationCalendar!=""){
VacationCalendar_1.Click();
 WorkspaceUtils.SearchByValue(VacationCalendar_1,"Vacation Calendar",VacationCalendar,"Vacation Calendar");
     }else{ 
  ValidationUtils.verify(false,true,"Vacation Calendar is Needed to Create a Employee");
}

    
var WeekCalendarNo_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McValuePickerWidget", "", 2);
if(WeekCalendarNo!=""){
 WeekCalendarNo_1.Click();
    WorkspaceUtils.SearchByValue(WeekCalendarNo_1,"Week Calendar",WeekCalendarNo,"WeekCalendar No");

   }else{ 
  ValidationUtils.verify(false,true,"Week Calendar No is Needed to Create a Employee");
}

  
//if(Employee_detail[19]!=""){
//var MinimumWorkingHours = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2);
//MinimumWorkingHours.Click();
//Sys.Process("Maconomy").Refresh();
//WorkspaceUtils.DropDownList(Employee_detail[19]);
//
//}
//  
//if(Employee_detail[20]!=""){
//var Max_Working_Hrs_per_day = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
//Max_Working_Hrs_per_day.setText(Employee_detail[20]);
//}
//  
//if(Employee_detail[21]!=""){
//var WorkingHours_Monday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
//WorkingHours_Monday.setText(Employee_detail[21]);
//}
//  
//if(Employee_detail[22]!=""){
//var WorkingHours_Tuesday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
//WorkingHours_Tuesday.setText(Employee_detail[22]);
//}
//  
//if(Employee_detail[23]!=""){
//var WorkingHours_Wednesday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2);
//WorkingHours_Wednesday.setText(Employee_detail[23]);
//}
//  
//if(Employee_detail[24]!=""){
//var WorkingHours_Thursday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
//WorkingHours_Thursday.setText(Employee_detail[24]);
//}
//  
//if(Employee_detail[25]!=""){
//var WorkingHours_Friday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
//WorkingHours_Friday.setText(Employee_detail[25]);
//}
//  
//if(Employee_detail[26]!=""){
//var WorkingHours_Saturday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 12).SWTObject("McTextWidget", "", 2);
//WorkingHours_Saturday.setText(Employee_detail[26]);
//}
//  
//if(Employee_detail[27]!=""){
//var WorkingHours_Sunday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 13).SWTObject("McTextWidget", "", 2);
//Max_Working_Hrs_per_day.setText(Employee_detail[27]);
//}
//  
//if(Employee_detail[28]!=""){
//var CostPerHour = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 2); 
//CostPerHour.setText(Employee_detail[28]);
//}
//  
//if(Employee_detail[29]!=""){
//var CreateEmployeeVendorAccount = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("McPopupPickerWidget", "", 2);
// CreateEmployeeVendorAccount.Click();
//
//Sys.Process("Maconomy").Refresh();
//var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
//var Add_Visible7 = true;
//while(Add_Visible7){
//if(list.isEnabled()){
//Add_Visible7 = false;
//    for(var i=list.getItemCount()-1;i>=0;i--){ 
//      if(list.getItem(i).getText_2(0)!=null){ 
//        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Employee_detail[29]){ 
//          list.Keys("[Enter]");
//
//          Delay(5000);
//          break;
//        }else{ 
//          list.Keys("[Up]");
//        }
//          
//      }else{ 
//        list.Keys("[Up]");
//      }
//    }
//}
//}
//
//
//  Delay(5000);
//}


var CreateUser_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
if(CreateUser_1.getSelection()){ 
  CreateUser_1.Click();
   ValidationUtils.verify(true,true,"Create user is Unselected");
  }
else{ 
  ReportUtils.logStep("INFO","Create User Check Box is Already UnSelected");
}
    
    
Delay(2000); 
var submit = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Submit");
Sys.HighlightObject(submit);
submit.HoverMouse();
ReportUtils.logStep_Screenshot();
submit.Click();
Delay(6000); 
  
var label1 = Sys.Process("Maconomy").WaitSWTObject("Shell", "Employees - Employee",1,10000).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}

}


function Employess(){ 
var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1)
.SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Employees");
All_Emp.Click();
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(comapany);
//  firstCell.Keys("1707");
firstCell.Keys("[Tab][Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//  job.setText("GAIL C COUTINHO")
job.setText(FullName + " "+STIME);
Delay(6000);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(2).OleValue.toString().trim()==(FullName + " "+STIME)){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
    
    
    
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Employee Created is available in system");
  
  
  
if(flag){ 
closefilter.Click();
Delay(5000);
empNumber  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 3).getText().toString().trim();
//Log.Message(empNumber);
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Employee No",EnvParams.Opco,"Data Management",empNumber)
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//if(ImageRepository.ImageSet.Forward.Exists()){
//ImageRepository.ImageSet.Forward.Click();// GL
//}

var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).WaitSWTObject("TabControl", "",1,60000);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
Delay(2000);
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 6,60000);
Delay(1000);
All_approver.Click();
Delay(3000);
ReportUtils.logStep_Screenshot();
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
//  Log.Message(approver_table.getItemCount());
var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   approvers="";
   if(approver_table.getItem(z).getText_2(8)!="Approved"){
   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   ReportUtils.logStep("INFO","Employee Approver level : " +z+ " Approver :" +approvers);
//   Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
   Approve_Level[y] = comapany+"*"+empNumber+"*"+approvers;
   Log.Message(Approve_Level[y])
   y++;
   }
}
}
var info_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
info_Bar.Click();
Delay(4000);

ImageRepository.ImageSet.Forward.Click();
Delay(4000);
var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
level = 1;
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
  
  Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}



Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}


Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}

Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}

Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
ValidationUtils.verify(true,true,"Created Employee is Approved by :"+Project_manager)
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(comapany+" - "+empNumber+" - Approver :"+Project_manager);
}
//}

if(Approve_Level.length==1){
var Employee_Vendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Employee_Vendor.Click();
Delay(7000);

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(comapany);
//firstCell.Keys("1707");
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);

job.setText(FullName + " "+STIME);
//job.setText("GAIL C COUTINHO");
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
Delay(6000);
  
  
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(4).OleValue.toString().trim()==(FullName + " "+STIME)){ 
//      if(table.getItem(v).getText_2(4).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;
    break;
  }else{ 
    table.Keys("[Down]");
  }
}
    
    
    
ReportUtils.logStep_Screenshot();

ValidationUtils.verify(flag,true,"Created Employee Vendor is available in system");
  
  
  
if(flag){ 
closeFilter.Click();
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
Delay(2000);
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
Delay(2000);
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Delay(1000);
All_approver.Click();


Delay(3000);
ReportUtils.logStep_Screenshot();
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  Log.Message(approver_table.getItemCount());
var y=0;
//for(var z=0;z<approver_table.getItemCount();z++){ 
//   approvers="";
//   if(approver_table.getItem(z).getText_2(8)!="Approved"){
////       approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   Log.Message("Employee Vendor Approver level : " +z+ " Approver :" +approvers);
//  ReportUtils.logStep("INFO","Global Employee Vendor Approver level : " +z+ " Approver :" +approvers);
////   Emp_Vendor_Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
//   Emp_Vendor_Approve_Level[y] = approvers;
//       Log.Message(Emp_Vendor_Approve_Level[y]);
//   y++;
//   }
//}


}
var infobar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
infobar.Click();
Delay(3000);
ImageRepository.ImageSet.Forward.Click();  




var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
Delay(2000);
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
Delay(2000);
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Delay(1000);
All_approver.Click();


Delay(3000);
ReportUtils.logStep_Screenshot();
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  Log.Message(approver_table.getItemCount());
var y=0;
//for(var z=0;z<approver_table.getItemCount();z++){ 
//   approvers="";
//   if(approver_table.getItem(z).getText_2(8)!="Approved"){
////       approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   Log.Message("Employee Vendor Approver level : " +z+ " Approver :" +approvers);
//   ReportUtils.logStep("INFO","Company Employee Vendor Approver level : " +z+ " Approver :" +approvers);
////   Emp_Vendor_Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
//   Emp_Vendor_Approve_Level[y] = approvers;
//       Log.Message(Emp_Vendor_Approve_Level[y]);
//   y++;
//   }
//}


}
var infobar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
infobar.Click();
Delay(3000);
ImageRepository.ImageSet.Forward.Click();  

}

}
}
}
}
}


function FinalApproveEmployess(comapany,empNumber,userNmae,apvLvl){ 
//function FinalApproveEmployess(){ 
//var ApproveInfo = [];
//ApproveInfo[0] = "1"
//var comapany = "1307";
//var empNumber = "130710098"
//var userNmae = "1307 Senoir Finance"
//var apvLvl = "0"

//var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1)
//.SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Employees");
//All_Emp.Click();
aqUtils.Delay(5000, Indicator.Text);
// Delay(5000) 
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
//Delay(5000)
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);
//Delay(3000);
} 

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(comapany);
//  firstCell.Keys("1707");
firstCell.Keys("[Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  job.setText(empNumber)
//job.setText(FullName + " "+STIME);
Delay(6000);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==empNumber){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
    
    
ReportUtils.logStep_Screenshot();  

ValidationUtils.verify(flag,true,"Employee Created is available in system");
  
  
  
if(flag){ 
closefilter.Click();
Delay(5000);
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
  
  Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}



Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}


Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}

Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}

Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
ValidationUtils.verify(true,true,"Created Employee is Approved by :"+userNmae)
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(comapany+" - "+empNumber+" - Approver :"+userNmae);
}
Delay(4000);
if(apvLvl==(ApproveInfo.length-1)){
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//if(ImageRepository.ImageSet.Forward.Exists()){
//ImageRepository.ImageSet.Forward.Click();// GL
//}

var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).WaitSWTObject("TabControl", "",1,60000);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
Delay(2000);
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 6,60000);
Delay(1000);
All_approver.Click();
Delay(3000);

ReportUtils.logStep_Screenshot();
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
//  Log.Message(approver_table.getItemCount());
var y=0;
//for(var z=0;z<approver_table.getItemCount();z++){ 
//   approvers="";
//   if(approver_table.getItem(z).getText_2(8)!="Approved"){
//   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   ReportUtils.logStep("INFO","Employee Approver level : " +z+ " Approver :" +approvers);
////   Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
//   Approve_Level[y] = comapany+"*"+empNumber+"*"+approvers;
//   Log.Message(Approve_Level[y])
//   y++;
//   }
//}
}
var info_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
info_Bar.Click();
Delay(4000);

ImageRepository.ImageSet.Forward.Click();
Delay(4000);
//var OpCo1 = EnvParams.Opco;
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
//var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
//if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
//level = 1;
//var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
//Approve.Click();
//Delay(4000);
//}

//if(Approve_Level.length==1){

//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var Employee_Vendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Employee_Vendor.HoverMouse();
ReportUtils.logStep_Screenshot();
Employee_Vendor.Click();
Delay(7000);

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(comapany);
//firstCell.Keys("1707");
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
//firstCell.Keys("[Tab]");
//Delay(1000);
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");

//job.setText(FullName + " "+STIME);
job.setText(empNumber);
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
Delay(6000);
  
  
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(3).OleValue.toString().trim()==empNumber){ 
//      if(table.getItem(v).getText_2(4).OleValue.toString().trim()=="GAIL C COUTINHO"){
    flag=true;
    break;
  }else{ 
    table.Keys("[Down]");
  }
}
    
    
    
    

ValidationUtils.verify(flag,true,"Created Employee Vendor is available in system");
  
  
  
if(flag){ 
closeFilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closeFilter.Click();
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
Delay(2000);
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot(); 
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
Delay(2000);
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Delay(1000);
All_approver.Click();


Delay(3000);
ReportUtils.logStep_Screenshot(); 
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  Log.Message(approver_table.getItemCount());
var y=0;
//for(var z=0;z<approver_table.getItemCount();z++){ 
//   approvers="";
//   if(approver_table.getItem(z).getText_2(8)!="Approved"){
////       approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   Log.Message("Employee Vendor Approver level : " +z+ " Approver :" +approvers);
//  ReportUtils.logStep("INFO","Global Employee Vendor Approver level : " +z+ " Approver :" +approvers);
////   Emp_Vendor_Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
//   Emp_Vendor_Approve_Level[y] = approvers;
//       Log.Message(Emp_Vendor_Approve_Level[y]);
//   y++;
//   }
//}


}
var infobar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
infobar.Click();
Delay(3000);
ImageRepository.ImageSet.Forward.Click();  




var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
Delay(2000);
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot(); 
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
Delay(2000);
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Delay(1000);
All_approver.Click();


Delay(3000);

ReportUtils.logStep_Screenshot(); 
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  Log.Message(approver_table.getItemCount());
var y=0;
//for(var z=0;z<approver_table.getItemCount();z++){ 
//   approvers="";
//   if(approver_table.getItem(z).getText_2(8)!="Approved"){
////       approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   Log.Message("Employee Vendor Approver level : " +z+ " Approver :" +approvers);
//   ReportUtils.logStep("INFO","Company Employee Vendor Approver level : " +z+ " Approver :" +approvers);
////   Emp_Vendor_Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
//   Emp_Vendor_Approve_Level[y] = approvers;
//       Log.Message(Emp_Vendor_Approve_Level[y]);
//   y++;
//   }
//}


}
var infobar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
infobar.Click();
Delay(3000);
ImageRepository.ImageSet.Forward.Click();  

}

//}
}
}
}
}

















function ApproveEmployee(){ 
  
var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Employees");
All_Emp.Click();
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(temp_user[0]);
//firstCell.Keys("1707");
firstCell.Keys("[Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
job.setText(temp_user[1]+ ""+STIME);
//job.setText(temp_user[1]);
Delay(6000);


  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(temp_user[1]+" "+STIME)){ 
//      if(table.getItem(v).getText_2(1).OleValue.toString().trim()==temp_user[1]){ 
      flag=true;
      break;
    }else{ 
    table.Keys("[Down]");
  }
      
  }




if(flag){ 
closefilter.Click();
Delay(5000);

var approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8)
    Sys.HighlightObject(approve);
    approve.Click();
}
  
}


function ApproveEmployeeVendor(){ 
//WorkspaceUtils.closeAllWorkspaces();
//
//goToMenu();

Delay(3000);
var Emp_vendor =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
Emp_vendor.Click();
Delay(7000);
var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstcell.setText(temp_user[0]);
//firstcell.Keys("1707");
firstcell.Keys("[Tab]");
Delay(1000);
firstcell.Keys("[Tab]");
Delay(1000);
firstcell.Keys("[Tab]");
Delay(1000);
firstcell.Keys("[Tab]");
Delay(1000);
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//job.setText(temp_user[1]);
job.setText(temp_user[1]+ " "+STIME);
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
Delay(6000);


  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(4).OleValue.toString().trim()==(temp_user[1]+" "+STIME)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==temp_user[1]){ 
      flag=true;
      break;
    }else{ 
    table.Keys("[Down]");
  }
      
  }




if(flag){ 
closeFilter.Click();
Delay(4000);
var approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7)
Sys.HighlightObject(approve);
approve.Click();
}
}





function getExcel(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf(",")!=-1){
     var excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
}

function getExcelData(rowidentifier,column) { 
var temp = ""
//var array = "Validate_EmployeeCategories";
//var Opco = "1307"
var excelData = [];
//Log.Message("Execution completed,sending result to excel book , FileName:"+excelName+"sheetname:"+sheet);
  var app = Sys.OleObject("Excel.Application");
//  app.Visible = "True";
  var curArrayVals = [];  
  var book = app.Workbooks.Open(workBook);
  var sheet = book.Sheets.Item(sheetName);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;
//  Log.Message(columnCount);
//  Log.Message(rowCount);
  var arrays={};
  var idx =0;
  var col =0;
  var row = 0;
  for(var k = 1; k<=columnCount;k++){
  if(sheet.Cells.Item(1, k).Text.toString().trim().toUpperCase()==column.toUpperCase()){
  col = k;
  }
  }
  var rowStatus = false;
  for(var k = 1; k<=rowCount;k++){
  if(sheet.Cells.Item(k, 1).Text.toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
  row = k;
  rowStatus = true;
  }
  }
  
  if(rowStatus){ 
   temp = sheet.Cells.Item(row,  col).Text;
//   Log.Message(temp)
  }
  
  
// book.Save();
 app.Quit();
 
 
 if(temp.indexOf(",")!=-1){ 
//       Log.Message(temp)
      excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     

 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);

 return excelData;
}



function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf("*")!=-1){
     var excelData =  temp.split("*");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
}





function CredentialLogin(){ 
// var Credentials = [];
// Credentials[0] = "1307*1307200357*1307 Finance (13079505)*OpCo - Billers";
// Credentials[1] = "1307*1307200357*Chinese Manager 2 (120110071)*Chinese Employee 1 (130710040)";
// Credentials[2] = "1307*1307200357*Central Team - Client Management*Central Team - Vendor Management";
// 
// var workBook = "H:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\China\\DS_SYSTEST_EN.xls";
// var sheetName = "Agency Users";
// var sheetName = "SSC Users";
//Central Team - Vendor Management
//"1307*1307200357*Central Team - Client Management*SSC - Expense Cashiers"

for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
//    Log.Message(Cred[j])
//Log.Message(j)
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
//     var workBook = "H:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\China\\DS_SYSTEST_EN.xls";
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 
//    var workBook = "H:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\China\\DS_SYSTEST_EN.xls";
    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }
  else{ 
   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
    if(UserN){ 
      goToHR();
      UserN = false;
    }
    temp = searchNumber(Eno);
  }
//  Log.Message(temp)
  if(temp.length!=0){
    temp = temp+"*"+j;
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
  break;
  }
  }
  if((temp=="")||(temp==null))
  Log.Error("User Name is Not available for level :"+i);
//  Log.Message("Logins :"+temp);
}
WorkspaceUtils.closeAllWorkspaces();

// ExcelUtils.setExcelName(workBook, sheetName, true);
//
// Cred[2] = ExcelUtils.SSCLogin(Cred[2],"Username");
// Cred[3] = ExcelUtils.SSCLogin(Cred[3],"Username");

}


function todo(lvl){ 
    var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);
//  Delay(3000);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
//  Delay(2000);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
//  var refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
//  refresh.Click();
  
  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();

  
  
  aqUtils.Delay(15000, Indicator.Text);
//  Delay(15000);
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");

if(lvl==3){
Client_Managt.ClickItem("|Approve Employee (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Employee (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Employee (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Employee (*)");
}
break;
}
}
}



}


