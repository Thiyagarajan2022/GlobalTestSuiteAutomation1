﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var emp_info = "Employee_EmployeeVendor_User_Cr";
var emp_info = "New EEU";
var Credential = "userRoles";
var login_satuts = true;
var LoginEmp = [];
var HRData = [];
var temp_user = [];
var Employee_detail = [];
var Approve_Level = [];
var UserLevel = [];
var User_Login = [];
var Emp_Vendor_Approve_Level = [];
var approvers;
var STIME="";
function goToMenu(){ 
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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


if(ImageRepository.ImageSet.Emp_1.Exists()){
ImageRepository.ImageSet.Emp_1.DblClick();// GL
}
else if(ImageRepository.ImageSet.Emp_2.Exists()){
ImageRepository.ImageSet.Emp_2.DblClick();// GL
}
else if(ImageRepository.ImageSet.Emp_3.Exists()){
ImageRepository.ImageSet.Emp_3.DblClick();// GL
}
Delay(3000);

}


function goToEmployees(){ 

//  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
//  employees.DblClickItem("|Employees");
var Add_Visible0 = true;
var New_Employee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 3,60000);
while(Add_Visible0){
if(New_Employee.isEnabled()){
New_Employee.Click();
Add_Visible0 = false;
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
  
if(Employee_detail[0]!=""){
var FullName = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
FullName.Click();
FullName.setText(Employee_detail[0] + " "+STIME );
}else{ 
  ValidationUtils.verify(false,true,"Employee Name is Needed to Create a Employee");
}
  
if(Employee_detail[1]!=""){
var Gender = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
Gender.Click();
Gender.Keys(Employee_detail[1]);
Delay(5000);
  

}

var Country = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2)
if(Employee_detail[2]!=""){
Country.Click();

Sys.Process("Maconomy").Refresh();
WorkspaceUtils.DropDownList(Employee_detail[2])
/*
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible2 = true;
while(Add_Visible2){
if(list.isEnabled()){
Add_Visible2 = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Employee_detail[2]){

          Delay(2000);
          list.Keys("[Enter]");
          Delay(5000);
          break;
        }else{ 
          list.Keys("[Down]");
        }
          
      }else{ 
        list.Keys("[Down]");
      }
    }
}
}
*/
}else{ 
  ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
}
  
  
  
var Company = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Employee_detail[3]!=""){
Company.Click();
WorkspaceUtils.SearchByValue(Company,"Company",Employee_detail[3]);
/*   Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(Employee_detail[3]);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  Log.Message(table.getItemCount());
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Employee_detail[3]){
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
       }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        Company.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    Company.setText("");
  }
  */
}else{ 
  ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
}

  

  
//var AccessLevel = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McValuePickerWidget", "", 2)
//if(Employee_detail[4]!=""){
//AccessLevel.Click();
//WorkspaceUtils.SearchByValue(AccessLevel,"Option",Employee_detail[4]);
//}else{ 
//  ValidationUtils.verify(false,true,"Access Level is Needed to Create a Employee");
//}

if(Employee_detail[4]!=""){
var DateEmployed = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2)
WorkspaceUtils.CalenderDateSelection(DateEmployed,Employee_detail[4])
//  DateEmployed.setText(Employee_detail[5]);
}else{ 
  ValidationUtils.verify(false,true,"Date Employed is Needed to Create a Employee");
}

if(Employee_detail[5]!=""){
var TerminateDate = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2)
WorkspaceUtils.CalenderDateSelection(TerminateDate,Employee_detail[5])
//  DateEmployed.setText(Employee_detail[5]);
}
  
if(Employee_detail[6]!=""){  
var Position =  Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
Position.setText(Employee_detail[6]);
}else{ 
  ValidationUtils.verify(false,true,"Position is Needed to Create a Employee");
}
  
if(Employee_detail[7]!=""){ 
var Email = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
Email.setText(Employee_detail[7]);
  
  
}else{ 
  ValidationUtils.verify(false,true,"Email Id is Needed to Create a Employee");
}

  
var ApproverGroup = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2) 
if(Employee_detail[8]!=null){
ApproverGroup.Click();

Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible3 = true;
 while(Add_Visible3){
if(list.isEnabled()){
Add_Visible3 = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Employee_detail[8]){ 
          Delay(1000);

            
          list.Keys("[Enter]");
          Delay(3000);
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
  
var EmploymentType = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2);
if(Employee_detail[9]!=""){
EmploymentType.Click();  
Delay(5000);
Sys.Process("Maconomy").Refresh(); 
WorkspaceUtils.DropDownList(Employee_detail[9]); 

 }else{ 
  ValidationUtils.verify(false,true,"Employment Type is Needed to Create a Employee");
} 
  
  
var EmployeeDepartment = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Employee_detail[10]!=""){
EmployeeDepartment.Click();
WorkspaceUtils.SearchByValue(EmployeeDepartment,"Local Specification 3",Employee_detail[10]);

 }else{ 
  ValidationUtils.verify(false,true,"Employment Type is Needed to Create a Employee");
} 

    
var EmployeeCostCentre = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Employee_detail[11]!=""){
EmployeeCostCentre.Click();
WorkspaceUtils.SearchByValue(EmployeeCostCentre,"Local Specification 5",Employee_detail[11]);

 }else{ 
  ValidationUtils.verify(false,true,"Employee Cost Centre is Needed to Create a Employee");
}

    
var Supervisor = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Employee_detail[12]!=""){
Supervisor.Click();
WorkspaceUtils.SearchByValue(Supervisor,"Employee",Employee_detail[12]);

     }else{ 
  ValidationUtils.verify(false,true,"Supervisor Centre is Needed to Create a Employee");
} 

    
    
var AbsenceApprover = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Employee_detail[13]!=""){
AbsenceApprover.Click();
WorkspaceUtils.SearchByValue(AbsenceApprover,"Employee",Employee_detail[13]);

     }else{ 
  ValidationUtils.verify(false,true,"Absence Approver is Needed to Create a Employee");
}
    
    
var Secretary =  Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Employee_detail[14]!=""){
Secretary.Click();
WorkspaceUtils.SearchByValue(Secretary,"Employee",Employee_detail[14]);

}
  Delay(2000);
  var next = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >")
 Sys.HighlightObject(next);
 next.Click();


}
}
}


function Employee_Information1(){ 
Delay(2000);
var Role = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
if(Employee_detail[15]!=""){
 Role.Click();
 WorkspaceUtils.SearchByValue(Role,"Employee Category",Employee_detail[15]);
     }else{ 
  ValidationUtils.verify(false,true,"Role is Needed to Create a Employee Vendor");
} 

    
if(Employee_detail[16]!=""){
var MustUseTimeSheets = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
MustUseTimeSheets.Click();

Sys.Process("Maconomy").Refresh();
WorkspaceUtils.DropDownList(Employee_detail[16]);
}
  
var VacationCalendar = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McValuePickerWidget", "", 2);
if(Employee_detail[17]!=""){
VacationCalendar.Click();
 WorkspaceUtils.SearchByValue(VacationCalendar,"Vacation Calendar",Employee_detail[17]);
     }else{ 
  ValidationUtils.verify(false,true,"Vacation Calendar is Needed to Create a Employee Vendor");
}

    
var WeekCalendarNo = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McValuePickerWidget", "", 2);
if(Employee_detail[18]!=""){
 WeekCalendarNo.Click();
    WorkspaceUtils.SearchByValue(WeekCalendarNo,"Week Calendar",Employee_detail[18]);

   }else{ 
  ValidationUtils.verify(false,true,"Week Calendar No is Needed to Create a Employee Vendor");
}

  
if(Employee_detail[19]!=""){
var MinimumWorkingHours = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2);
MinimumWorkingHours.Click();
Sys.Process("Maconomy").Refresh();
WorkspaceUtils.DropDownList(Employee_detail[19]);

}
  
if(Employee_detail[20]!=""){
var Max_Working_Hrs_per_day = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
Max_Working_Hrs_per_day.setText(Employee_detail[20]);
}
  
if(Employee_detail[21]!=""){
var WorkingHours_Monday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
WorkingHours_Monday.setText(Employee_detail[21]);
}
  
if(Employee_detail[22]!=""){
var WorkingHours_Tuesday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
WorkingHours_Tuesday.setText(Employee_detail[22]);
}
  
if(Employee_detail[23]!=""){
var WorkingHours_Wednesday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2);
WorkingHours_Wednesday.setText(Employee_detail[23]);
}
  
if(Employee_detail[24]!=""){
var WorkingHours_Thursday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
WorkingHours_Thursday.setText(Employee_detail[24]);
}
  
if(Employee_detail[25]!=""){
var WorkingHours_Friday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
WorkingHours_Friday.setText(Employee_detail[25]);
}
  
if(Employee_detail[26]!=""){
var WorkingHours_Saturday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 12).SWTObject("McTextWidget", "", 2);
WorkingHours_Saturday.setText(Employee_detail[26]);
}
  
if(Employee_detail[27]!=""){
var WorkingHours_Sunday = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 13).SWTObject("McTextWidget", "", 2);
Max_Working_Hrs_per_day.setText(Employee_detail[27]);
}
  
if(Employee_detail[28]!=""){
var CostPerHour = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 2); 
CostPerHour.setText(Employee_detail[28]);
}
  
if(Employee_detail[29]!=""){
var CreateEmployeeVendorAccount = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("McPopupPickerWidget", "", 2);
 CreateEmployeeVendorAccount.Click();

Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible7 = true;
while(Add_Visible7){
if(list.isEnabled()){
Add_Visible7 = false;
    for(var i=list.getItemCount()-1;i>=0;i--){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Employee_detail[29]){ 
          list.Keys("[Enter]");

          Delay(5000);
          break;
        }else{ 
          list.Keys("[Up]");
        }
          
      }else{ 
        list.Keys("[Up]");
      }
    }
}
}


  Delay(5000);
}


var checkBox = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
if(!checkBox.getSelection()){ 
  checkBox.Click();
    Log.Message("Create User checkBox is Clicked")
  }
    
    
Delay(2000); 
var next = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
next.Click();
Delay(4000); 



}


function user(){ 

var user_type = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
if(Employee_detail[30]!=""){
user_type.Click();
 WorkspaceUtils.SearchByValue(user_type,"User Type",Employee_detail[30]);
     }else{ 
  ValidationUtils.verify(false,true,"User Type is Needed to Create a User");
}

var user_name = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
if(Employee_detail[31]!=""){
user_name.setText(Employee_detail[31]+" "+STIME);
Delay(2000);
}else{ 
  ValidationUtils.verify(false,true,"User Name is Needed to Create a User");
}

var valid_period_from = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McDatePickerWidget", "", 2);
if(Employee_detail[32]!=""){ 
WorkspaceUtils.CalenderDateSelection(valid_period_from,Employee_detail[32]);
}else{ 
  ValidationUtils.verify(false,true,"Valid Perform is Needed to Create a User");
}

var valid_period_to = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McDatePickerWidget", "", 4);
if(Employee_detail[33]!=""){ 
 WorkspaceUtils.CalenderDateSelection(valid_period_to,Employee_detail[33]); 
}else{ 
  ValidationUtils.verify(false,true,"Valid Perform is Needed to Create a User");
}

var user_Access_Level = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McValuePickerWidget", "", 2);
user_Access_Level.Click();
if((user_Access_Level.getText()!=Employee_detail[34]) &&(Employee_detail[34]!="")){ 
  WorkspaceUtils.SearchByValue(user_Access_Level,"Option",Employee_detail[34]); 
}
var submitUser = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2);
if((Employee_detail[35]!="")&&(submitUser.getText()!=Employee_detail[35])){ 
submitUser.Click();
Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible8 = true;
while(Add_Visible8){
if(list.isEnabled()){
Add_Visible8 = false;
    for(var i=list.getItemCount()-1;i>=0;i--){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Employee_detail[35]){ 
          list.Keys("[Enter]");

          Delay(5000);
          break;
        }else{ 
          list.Keys("[Up]");
        }
          
      }else{ 
        list.Keys("[Up]");
      }
    }
}
}


  Delay(5000);

}


var create = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
create.Click();
Delay(6000); 
  
var label1 = Sys.Process("Maconomy").WaitSWTObject("Shell", "Employees - Employee",1,10000).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.Click();
Delay(8000); 
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.Click();
Delay(8000); 
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.Click();
Delay(8000); 
var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
Ok.Click();
  
//  var idex = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").Index
//  for(var i=1;i<=idex;i++){
//  var label = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", i).SWTObject("Label", "*").WndCaption;
//  if(label.indexOf("is added to the membership list of category")!=-1){ 
//  
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", i).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  Delay(6000);
//  }
//  else if(label.indexOf("Employee, Vendor, Company Vendor and User accounts have been created and submitted successfully")!=-1){ 
//  
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", i).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  Delay(6000);
//  }else if(label.indexOf("Password: CORE@WPP123")!=-1){ 
//  
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", i).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  Delay(6000);
//  }else if(label.indexOf("The expiry date is not within the validity period")!=-1){ 
//  
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", i).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  Delay(6000);
//  }
//   
//  } 
//  if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption == "Employees - Employee"){ 
//  var label = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Label", "*").WndCaption;
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  }
//  Delay(6000); 
//  
//  if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption == "Employees - Employee"){ 
//  var label = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Label", "*").WndCaption;
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  }
//  Delay(6000); 
//  
//  if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption == "Employees - Employee"){ 
//  var label = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Label", "*").WndCaption;
//  Log.Message(label);
//  var OkButton = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//  OkButton.Click();
//  }
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 1)
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 1).SWTObject("Label", "Employee 170710154 is added to the membership list of category EC1025")
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//  
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 1)
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 1).SWTObject("Label", "Employee, Vendor, Company Vendor and User accounts have been created and submitted successfully")
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//  
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 3)
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 3).SWTObject("Label", "Password: CORE@WPP123")
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 3).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//  
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 4)
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 4).SWTObject("Label", "The expiry date is not within the validity period")
//  Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee", 4).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//  var label = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Label", "*").getText();
//  Log.Message(label);
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//  Ok.Click();
//
//  Delay(8000); 
//
//var label1 = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Label", "*");
//Log.Message(label);
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Employee").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//Ok.Click();
}




function Employess(){ 
var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1)
.SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Employees");
All_Emp.Click();
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(Employee_detail[3]);
//  firstCell.Keys("1707");
firstCell.Keys("[Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//  job.setText("Employee Sample 27December2018 11:46:52")
job.setText(Employee_detail[0]+ " "+STIME);
Delay(6000);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(Employee_detail[0]+" "+STIME)){ 
//if(table.getItem(v).getText_2(1).OleValue.toString().trim()=="Employee Sample 27December2018 11:46:52"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
    
    
    

ValidationUtils.verify(flag,true,"Employee Created is available in system");
  
  
  
if(flag){ 
closefilter.Click();
Delay(5000);
var empNumber  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "")
.SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2).getText();
Log.Message(empNumber);
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).WaitSWTObject("TabControl", "",1,60000);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
Delay(2000);
Add_Visible8 = false;
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 7,60000);
Delay(1000);
All_approver.Click();
Delay(3000);
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
//  Log.Message(approver_table.getItemCount());
var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   approvers="";
   if(approver_table.getItem(z).getText_2(8)!="Approved"){
   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   ReportUtils.logStep("INFO","Employee Approver level : " +z+ " Approver :" +approvers);
   Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
   Log.Message(Approve_Level[y])
   y++;
   }
}
}
var info_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
info_Bar.Click();
Delay(4000);

//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//var users = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
//users.Click();
//var username = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//username.Click();
//username.setText(Employee_detail[31]+ " "+STIME);
////username.Keys("MMuthu 27December2018 11:46:52");
//var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
//Sys.HighlightObject(table)
//flag = false;
//  for(var v=0;v<table.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(Employee_detail[31]+" "+STIME)){ 
////if(table.getItem(v).getText_2(0).OleValue.toString().trim()=="MMuthu 27December2018 11:46:52"){
//      flag=true;    
//      break;
//    }
//    else{ 
//      table.Keys("[Down]");
//    }
//  }
// ValidationUtils.verify(flag,true,"User Created is available in system");
//if(flag){
//var Cf = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//Cf.Click();
//Delay(4000);
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
//approve_Bar.Click();
//Delay(3000);
//var allAprove = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
//allAprove.Click();
//Delay(3000);
//var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
//Sys.HighlightObject(approver_table);
//var y=0;
//for(var z=0;z<approver_table.getItemCount();z++){ 
//     approvers="";
//     if(approver_table.getItem(z).getText_2(8)!="Approved"){
//     approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//     Log.Message("User Approver level : " +z+ " Approver :" +approvers);
//
//     UserLevel[y] = Employee_detail[31]+"*"+empNumber+"*"+approvers;
//     y++;
//     }
//  }
//}
//
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//Delay(5000);
//Delay(3000);
var Employee_Vendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Employee_Vendor.Click();
Delay(7000);

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(Employee_detail[3]);
//firstCell.Keys("1707");
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
firstCell.Keys("[Tab]");
Delay(1000);
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);

job.setText(Employee_detail[0]+ " "+STIME);
//job.setText("Employee Sample 27December2018 11:46:52");
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
Delay(6000);
  
  
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(4).OleValue.toString().trim()==(Employee_detail[0]+" "+STIME)){ 
//      if(table.getItem(v).getText_2(4).OleValue.toString().trim()=="Employee Sample 27December2018 11:46:52"){
    flag=true;
    break;
  }else{ 
    table.Keys("[Down]");
  }
}
    
    
    
    

ValidationUtils.verify(flag,true,"Employee Vendor Created is available in system");
  
  
  
if(flag){ 
closeFilter.Click();
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();

var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).WaitSWTObject("TabControl", "",1,60000);
if(approve_Bar.isEnabled()){
Delay(2000);
  
approve_Bar.Click();
Delay(2000);
ImageRepository.ImageSet.Maximize.Click();
Delay(2000);
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 5,60000);
Delay(1000);
All_approver.Click();


Delay(3000);
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  Log.Message(approver_table.getItemCount());
var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   approvers="";
   if(approver_table.getItem(z).getText_2(8)!="Approved"){
//       approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   Log.Message("Employee Vendor Approver level : " +z+ " Approver :" +approvers);

   Emp_Vendor_Approve_Level[y] = Employee_detail[3]+"*"+Employee_detail[0]+"*"+approvers;
       Log.Message(Emp_Vendor_Approve_Level[y]);
   y++;
   }
}
}
  
}
}
}

WorkspaceUtils.closeAllWorkspaces();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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
 
if(ImageRepository.ImageSet.User1.Exists()){
ImageRepository.ImageSet.User1.DblClick();// GL
}
else if(ImageRepository.ImageSet.User3.Exists()){
ImageRepository.ImageSet.User3.DblClick();// GL
}
else if(ImageRepository.ImageSet.User2.Exists()){
ImageRepository.ImageSet.User2.DblClick();// GL
}
Delay(3000);
var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click();
 
Delay(4000);
var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
firstcell.setText(Employee_detail[31]+ " "+STIME);
//  firstcell.Keys("MMuthu 27December2018 11:46:52");
firstcell.Keys("[Tab][Tab]");
Delay(2000);
var Employee_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
//  Employee_no.Keys("170710158");
Employee_no.Keys(empNumber);

  
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
 Delay(4000);
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(Employee_detail[31]+" "+STIME)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()=="MMuthu 27December2018 11:46:52"){ 
      flag=true;
      break;
    }else{ 
    table.Keys("[Down]");
  }
      
  }
  ValidationUtils.verify(flag,true,"User Created is available in system");
if(flag){ 
 var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  closefilter.Click();
  Delay(5000);
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var apprv = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
apprv.Click();
Delay(3000);
var allAprove = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
allAprove.Click();
Delay(3000);
var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
Sys.HighlightObject(approver_table);
var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
     approvers="";
     if(approver_table.getItem(z).getText_2(8)!="Approved"){
     approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
     Log.Message("User Approver level : " +z+ " Approver :" +approvers);

     UserLevel[y] = Employee_detail[31]+"*"+empNumber+"*"+approvers;
     Log.Message(UserLevel[y]);
     y++;
     }
  }
  }

 
 
  

}

function gettingLogIN() {
//Approve_Level = [];
//Approve_Level[0] = "1707*Employee Sample 27December2018 16:04:12*SACHINDRA P KARKERA (170710011)*";
//Emp_Vendor_Approve_Level[0] = "1707*Employee Sample 27December2018 16:04:12*SACHINDRA P KARKERA (170710011)*";
//UserLevel[0] = "MMuthu 27December2018 16:04:12*170710178*SACHINDRA P KARKERA (170710011)*"

Delay(3000);
goToHR();
Credentiallogin();
login_satuts = true;
if(login_satuts){
var Emp_Login = Login_Match(Approve_Level);
}
if(login_satuts){
var Emp_Vendor_Login = Login_Match(Emp_Vendor_Approve_Level);
}
if(login_satuts){
var User_Login = Login_Match(UserLevel);
}
var j=0;   
for(var i=0;i<Emp_Login.length;i++){ 
var temp = User_Login[i];
temp_user = [];
temp_user = temp.split("*");
   if((Emp_Login.length==Emp_Vendor_Login.length)&&(Emp_Login[i]==Emp_Vendor_Login[i])&&(Emp_Login.length==User_Login.length)&&(Emp_Login[i].indexOf(temp_user[2]))&&(Emp_Login[i].indexOf(temp_user[3]))){ 
     j++;
   }

}
if(j==Emp_Login.length){ 
RestMaconomy(Emp_Login,User_Login,1); 
}else{
RestMaconomy(Emp_Login,"","empOnly");
RestMaconomy(Emp_Vendor_Login,"","empVendor");
RestMaconomy(User_Login,"","User");
}
   
}


function RestMaconomy(UserPasswd,User_Login,state){ 
for(var i=0;i<UserPasswd.length;i++){
var temp = UserPasswd[i];
temp_user = [];
temp_user = temp.split("*");
var uname = temp_user[2]; 
var pwd = temp_user[3];
Rests(uname,pwd);
    
goToMenu();
if(state==1){ 
EMP_ONLY();
Employee1();
var temp = User_Login[i];
temp_user = [];
temp_user = temp.split("*");
userApprv();
}
if(state=="empOnly"){ 
EMP_ONLY();
}
if(state=="empVendor"){ 
Employee1();
}
if(state=="User"){ 
userApprv();
}
}
}

function Rests(uname,pwd){ 
Delay(5000);
  Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x52); //R 
 Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
 Sys.Desktop.KeyUp(0x52); //R
Delay(65000);
 var usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
var pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
var btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
usernameAddr.SetFocus();
usernameAddr.setText(uname);
pwdAddr.setText(pwd);
btnLogin.click();
Delay(10000);   
}





function EMP_ONLY(){ 
  
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


function Employee1(){ 
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

function userApprv(){ 
WorkspaceUtils.closeAllWorkspaces();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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
 
if(ImageRepository.ImageSet.User1.Exists()){
ImageRepository.ImageSet.User1.DblClick();// GL
}
else if(ImageRepository.ImageSet.User3.Exists()){
ImageRepository.ImageSet.User3.DblClick();// GL
}
else if(ImageRepository.ImageSet.User2.Exists()){
ImageRepository.ImageSet.User2.DblClick();// GL
}
Delay(3000);
var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click();
 
Delay(4000);
var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
firstcell.Keys(temp_user[0]+" "+STIME);
//  firstcell.Keys(temp_user[0]);
firstcell.Keys("[Tab][Tab]");
Delay(2000);
var Employee_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
//  Employee_no.Keys("170710158");
Employee_no.Keys(temp_user[1]);

  
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
 Delay(4000);
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(temp_user[0]+" "+STIME)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==temp_user[0]){ 
      flag=true;
      break;
    }else{ 
    table.Keys("[Down]");
  }
      
  }
  
if(flag){ 
 var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  closefilter.Click();
  Delay(5000);
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject( approve );
  approve.Click();
  Delay(5000);
  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  Ok.Click();
   WorkspaceUtils.closeAllWorkspaces();

  }

 
}










function goToHR(){ 
Delay(3000);
closeAllWorkspaces();
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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

var HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
HRitem.DblClickItem("|Users");
Delay(5000);
//var ActiveUser = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Users");
//ActiveUser.Click();
var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click();
Delay(5000);
var HRTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var z=0;
for(var i=0;i<HRTable.getItemCount();i++){ 
if(HRTable.getItem(i).getText(2)!=""){
HRData[z] = HRTable.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+HRTable.getItem(i).getText_2(2).OleValue.toString().trim()
//Log.Message(HRData[z]);
z++;

}
}

}


function Credentiallogin() {
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, "userRoles", false);
var id =0;
var colsList = [];

for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
 colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
}
 while (!DDT.CurrentDriver.EOF()) {
 var temp ="";
  for(var idx=0;idx<colsList.length;idx++){  
   if(xlDriver.Value(colsList[idx])!=null){
  temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
  }
  else{ 
    temp = temp+"*";
  }
  }
//      Log.Message(temp)
 LoginEmp[id]=temp;
 id++;     
 xlDriver.Next();
 }
 DDT.CloseDriver(xlDriver.Name);
}


function Login_Match(Approve_Level_1){ 
login_satuts = true;
Delay(3000);
var UserPasswd = [];
var z =0;
for(var i=0;i<Approve_Level_1.length;i++){ 
if((Approve_Level_1[i].indexOf("OpCo")!=-1) && (Employee_detail[3]!="")){
Approve_Level_1[i] = Approve_Level_1[i].replace(/OpCo/g,Employee_detail[3]);
}
// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
if(Approve_Level_1[i].indexOf("SSC - Biller")==-1){
Approve_Level_1[i] = Approve_Level_1[i].replace(/- Billers/g,"- Agency - Biller");
}

var tempLevel = Approve_Level_1[i].split("*");
ifGotIT = true;
for(var j=2;j<tempLevel.length;j++){ 

if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
var tempSplit = temp.split("*");

for(var k=0;k<LoginEmp.length;k++){
var A_temp = LoginEmp[k].split("*");
if(tempSplit[0]==A_temp[0]){ 
  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
 Log.Message(UserPasswd[z]);
 z++;
 ifGotIT = false;
 break;     
}else{ 
if(tempSplit[1]==A_temp[2]){ 
  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
 Log.Message(UserPasswd[z]);
 z++;
 ifGotIT = false;
 break;     
}     
}
    
}
if(!ifGotIT){ 
break;
}
}
  
if((tempLevel[j].indexOf("SSC -")!=-1) || (tempLevel[j].indexOf("Central Team -")!=-1)){ 
    
 if(tempLevel[j].indexOf("Central Team - Client Management")!=-1){ 
  temp2 = "Central Team - Client Account Management";
}
else if(tempLevel[j].indexOf("Central Team - Vendor Management")!=-1){ 
  temp2 = "Central Team - Vendor Account Management";
}
else if(tempLevel[j].indexOf("SSC - Expense Cashiers")!=-1){ 
  temp2 = "SSC - Cashier";
}
for(var k=0;k<LoginEmp.length;k++){
var A_temp = LoginEmp[k].split("*");  
if(tempLevel[j]==A_temp[1]){ 
 UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
 Log.Message(UserPasswd[z]);
 z++;
 ifGotIT = false;
 break;     
}
}  

if(!ifGotIT){ 
break;
}
}
  
  
  
if((tempLevel[j].indexOf(" (")==-1) && (tempLevel[j].indexOf(")")==-1) && 
(tempLevel[j].indexOf("SSC -")==-1) && (tempLevel[j].indexOf("Central Team -")==-1)){ 
    
for(var k=0;k<LoginEmp.length;k++){
var A_temp = LoginEmp[k].split("*");
if(A_temp[0]==tempLevel[j]){  // Better  to use level[j].indexOf(LoginArrays[k])
UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]; 
Log.Message(UserPasswd[z]);
 z++;
 ifGotIT = false;
 break;     
}
}
if(!ifGotIT){ 
break;
}
}
  
if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){

var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
var tempSplit = temp.split("*");

for(var k=0;k<HRData.length;k++){
var A_temp = HRData[k].split("*");
if(tempSplit[1]==A_temp[1]){ 
 UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
 Log.Message(UserPasswd[z]);
 z++;
 ifGotIT = false;
 break;     
}
    
}
if(!ifGotIT){ 
break;
}
}
 
  
}
if(ifGotIT){ 
Log.Warning("UserName and Password is Not Matched for Approver and Substitute :"+Approve_Level[i]);
login_satuts = false;
break;
}
  
}

return UserPasswd;
}



function excel(sheetName){ 
var Arrays = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

 for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
   colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
 }
   while (!DDT.CurrentDriver.EOF()) {
   var temp ="";
    for(var idx=1;idx<colsList.length;idx++){  
     if(xlDriver.Value(colsList[idx])!=null){
    temp = temp+xlDriver.Value(colsList[idx]).toString().trim();
    }
    else{ 
      temp = temp;
    }
    }
   Arrays[id]=temp;
   id++;     
   xlDriver.Next();
   }
   DDT.CloseDriver(xlDriver.Name);
return Arrays;
}


function CreateEmployeeUser(){ 
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create Employee test started::"+STIME);
goToMenu();  

goToEmployees();
Employee_detail = [];
Employee_detail = excel(emp_info)
Employee_Information();
Employee_Information1();
user();
//  Log.Message("Employee is Created");
Employess();
//  WorkspaceUtils.closeAllWorkspaces();
gettingLogIN();
WorkspaceUtils.closeAllWorkspaces();
}


