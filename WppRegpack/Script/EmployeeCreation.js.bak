//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart
//var excelName = EnvParams.getEnvironment();

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "EmployeeAsUser";
var FullName,Gender,Country,comapany,DateEmployed,Position,Email,ApproverGroup,EmploymentType,SalesEmployee,EmployeeDepartment_Name,EmployeeCostCentre_Name,Supervisor,AbsenceApprover,Role,VacationCalendar,WeekCalendarNo,CreateUser,UserType,ValidityPeriodFrom,ValidityPeriodTo,AccessLevel,Language = "";
var Language = "";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var empNumber = "";
 
var emp_info = "Employee_EmployeeVendor_User_Cr";
var emp_info = "New EEU";
var Credential = "userRoles";
var login_satuts = true;
var LoginEmp = [];
var HRData = [];
var temp_user = [];
var Employee_detail = [];

var UserLevel = [];
var User_Login = [];
var Emp_Vendor_Approve_Level = [];
var approvers;
var STIME="";
var Project_manager = "";
function goToMenu(){ 
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
  menuBar.DblClick();

if(ImageRepository.ImageSet.Human_Resource_1.Exists()){
ImageRepository.ImageSet.Human_Resource_1.Click();
}  
else if(ImageRepository.ImageSet.Human_Resource.Exists()){
ImageRepository.ImageSet.Human_Resource.Click();
}
else if(ImageRepository.ImageSet.HumanResource.Exists()){
ImageRepository.ImageSet.HumanResource.Click();
}
else if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();  
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

  aqUtils.Delay(3000, "Finding Employee");
//  Delay(3000);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
//  Delay(2000);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, "Finding Employee");
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees").OleValue.toString().trim());

}

}
//Delay(3000);
TextUtils.writeLog("Entering into Employees from Human Resources Menu");
}


function goToEmployees(){ 

//  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
//  employees.DblClickItem("|Employees");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var Add_Visible0 = true;
var New_Employee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 3,60000);
while(Add_Visible0){
if(New_Employee.isEnabled()){
New_Employee.HoverMouse();
ReportUtils.logStep_Screenshot();
New_Employee.Click();
Add_Visible0 = false;
}
}
 TextUtils.writeLog("New Employee is Clicked"); 
}



//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
FullName = ExcelUtils.getRowDatas("FullName",EnvParams.Opco)
if((FullName==null)||(FullName=="")){ 
ValidationUtils.verify(false,true,"FullName is Needed to Create a Employee");
}
Gender = ExcelUtils.getRowDatas("Gender",EnvParams.Opco)
if((Gender==null)||(Gender=="")){ 
ValidationUtils.verify(false,true,"Gender is Needed to Create a Employee");
}
Country = ExcelUtils.getRowDatas("Country",EnvParams.Opco)
if((Country==null)||(Country=="")){ 
ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
}

comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
Log.Message(comapany)
if((comapany==null)||(comapany=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Employee");
}
DateEmployed = ExcelUtils.getRowDatas("DateEmployed",EnvParams.Opco)
if((DateEmployed==null)||(DateEmployed=="")){ 
ValidationUtils.verify(false,true,"DateEmployed is Needed to Create a Employee");
}
Position = ExcelUtils.getRowDatas("Position",EnvParams.Opco)
if((Position==null)||(Position=="")){ 
ValidationUtils.verify(false,true,"Position is Needed to Create a Employee");
}
Email = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
if((Email==null)||(Email=="")){ 
ValidationUtils.verify(false,true,"Email is Needed to Create a Employee");
}
ApproverGroup = ExcelUtils.getRowDatas("ApproverGroup",EnvParams.Opco)

EmploymentType= ExcelUtils.getRowDatas("EmploymentType",EnvParams.Opco)
if((EmploymentType==null)||(EmploymentType=="")){ 
ValidationUtils.verify(false,true,"EmploymentType is Needed to Create a Employee");
}
SalesEmployee= ExcelUtils.getRowDatas("Sales Employee",EnvParams.Opco)
if((SalesEmployee==null)||(SalesEmployee=="")){ 
ValidationUtils.verify(false,true,"Sales Employee is Needed to Create a Employee");
}
EmployeeDepartment_Name= ExcelUtils.getRowDatas("EmployeeDepartment_Name",EnvParams.Opco)
if((EmployeeDepartment_Name==null)||(EmployeeDepartment_Name=="")){ 
ValidationUtils.verify(false,true,"EmployeeDepartment_Name is Needed to Create a Employee");
}
EmployeeCostCentre_Name= ExcelUtils.getRowDatas("EmployeeCostCentre_Name",EnvParams.Opco)
if((EmployeeCostCentre_Name==null)||(EmployeeCostCentre_Name=="")){ 
ValidationUtils.verify(false,true,"EmployeeCostCentre_Name is Needed to Create a Employee");
}
Supervisor= ExcelUtils.getRowDatas("Supervisor",EnvParams.Opco)
if((Supervisor==null)||(Supervisor=="")){ 
ValidationUtils.verify(false,true,"Supervisor is Needed to Create a Employee");
}
AbsenceApprover= ExcelUtils.getRowDatas("AbsenceApprover",EnvParams.Opco)
if((AbsenceApprover==null)||(AbsenceApprover=="")){ 
ValidationUtils.verify(false,true,"AbsenceApprover is Needed to Create a Employee");
}
Role= ExcelUtils.getRowDatas("Role",EnvParams.Opco)
if((Role==null)||(Role=="")){ 
ValidationUtils.verify(false,true,"Role is Needed to Create a Employee");
}
VacationCalendar= ExcelUtils.getRowDatas("Vacation Calendar",EnvParams.Opco)
if((VacationCalendar==null)||(VacationCalendar=="")){ 
ValidationUtils.verify(false,true,"Vacation Calendar is Needed to Create a Employee");
}
WeekCalendarNo= ExcelUtils.getRowDatas("Week Calendar No.",EnvParams.Opco)
if((WeekCalendarNo==null)||(WeekCalendarNo=="")){ 
ValidationUtils.verify(false,true,"Week Calendar No. is Needed to Create a Employee");
}

}


function EmployeeScreen1_Address(){ 
//Checking Labels in Job Create Wizard
Delay(4000);
Sys.Process("Maconomy").Refresh();
var FullName_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
if(FullName_1!="Full Name")
ValidationUtils.verify(false,true,"Full Name field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Full Name field is available in Maconomy for Employee Creation");
//var Gender_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//if(Gender_1!="Gender")
//ValidationUtils.verify(false,true,"Gender field is missing in Maconomy for Employee Creation");
//else
//ValidationUtils.verify(true,true,"Gender field is available in Maconomy for Employee Creation");
var Country_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Country_1!="Country")
ValidationUtils.verify(false,true,"Country field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Country field is available in Maconomy for Employee Creation");
var Company_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Company_1!="Company")
ValidationUtils.verify(false,true,"Company field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Company field is available in Maconomy for Employee Creation");
var DateEmployed_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(DateEmployed_1!="Date Employed")
ValidationUtils.verify(false,true,"Date Employed field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Date Employed field is available in Maconomy for Employee Creation");
var Position_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Position_1!="Position")
ValidationUtils.verify(false,true,"Position field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Position field is available in Maconomy for Employee Creation");
var Email_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Email_1!="E-mail")
ValidationUtils.verify(false,true,"E-mail field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"E-mail field is available in Maconomy for Employee Creation");
var ApproverGroup_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(ApproverGroup_1!="Approver Group")
ValidationUtils.verify(false,true,"Approver Group field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Approver Group field is available in Maconomy for Employee Creation");
var EmploymentType_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(EmploymentType_1!="Employment Type")
ValidationUtils.verify(false,true,"Employment Type field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Employment Type field is available in Maconomy for Employee Creation");
var SalesEmployee_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("McTextWidget", "").getText().OleValue.toString().trim()
if(SalesEmployee_1!="Sales Employee")
ValidationUtils.verify(false,true,"Sales Employee field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Sales Employee field is available in Maconomy for Employee Creation");
var EmployeeDepartment_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(EmployeeDepartment_1!="Employee Department")
ValidationUtils.verify(false,true,"Employee Department field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Employee Department field is available in Maconomy for Employee Creation");
var EmployeeCostCentre_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(EmployeeCostCentre_1!="Employee Cost Centre")
ValidationUtils.verify(false,true,"Employee Cost Centre field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Employee Cost Centre field is available in Maconomy for Employee Creation");
var Supervisor_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Supervisor_1!="Supervisor")
ValidationUtils.verify(false,true,"Supervisor field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Supervisor field is available in Maconomy for Employee Creation");
var AbsenceApprover_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(AbsenceApprover_1!="Absence Approver")
ValidationUtils.verify(false,true,"Absence Approver field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Absence Approver field is available in Maconomy for Employee Creation");

}


function EmployeeScreen2_Address(){ 
var Role_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Role_1!="Role")
ValidationUtils.verify(false,true,"Role field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Role field is available in Maconomy for Employee Creation");
var VacationCalendar_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(VacationCalendar_1!="Vacation Calendar")
ValidationUtils.verify(false,true,"Vacation Calendar field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Vacation Calendar field is available in Maconomy for Employee Creation");
var WeekCalendarNo_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(WeekCalendarNo_1!="Week Calendar No.")
ValidationUtils.verify(false,true,"WeekCalendar No field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"WeekCalendar No field is available in Maconomy for Employee Creation");
var CreateUser_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McPlainCheckboxView", "", 2).SWTObject("McTextWidget", "").getText().OleValue.toString().trim()
if(CreateUser_1!="Create User")
ValidationUtils.verify(false,true,"Create User field is missing in Maconomy for Employee Creation");
else
ValidationUtils.verify(true,true,"Create User field is available in Maconomy for Employee Creation");
}




function Employee_Information(){ 
Delay(5000);
var Name = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("McTextWidget", "", 2,60000);
var Add_Visible = true;
while(Add_Visible){
if(Name.isEnabled()){
Delay(2000);
Add_Visible = false;
  
if(FullName!=""){
var FullName_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
FullName_1.Click();
FullName_1.setText(FullName + " "+STIME );
ValidationUtils.verify(true,true,"Employee Name is entered in Maconomy");
}else{ 
  ValidationUtils.verify(false,true,"Employee Name is Needed to Create a Employee");
}
  
//if(Gender!=""){
//var Gender_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
//Gender_1.Click();
//Gender_1.Keys(Gender);
//Delay(5000);
//ValidationUtils.verify(true,true,"Gender is selected in Maconomy"); 
//
//}

var Country_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2)
if(Country!=""){
Country_1.Click();

Sys.Process("Maconomy").Refresh();
WorkspaceUtils.DropDownList(Country,"Country");
}else{ 
  ValidationUtils.verify(false,true,"Country is Needed to Create a Employee");
}
  
  

 
var Company_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
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
var DateEmployed_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2)
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
var Position_1 =  Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
Position_1.setText(Position);
ValidationUtils.verify(true,true,"Position is Entered in Maconomy"); 
}else{ 
  ValidationUtils.verify(false,true,"Position is Needed to Create a Employee");
}
  
if(Email!=""){ 
var Email_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
var Eml_split1 = Email.substring(0,Email.indexOf("@"));
var Eml_split2 = Email.substring(Email.indexOf("@"));
Eml_split1 = Eml_split1 +" "+STIME;
Eml_split1 = Eml_split1.replace(/[_: ]/g,"");
Email = Eml_split1+Eml_split2
Email_1.setText(Email);  
ValidationUtils.verify(true,true,"Email Id is Entered in Maconomy"); 
}else{ 
  ValidationUtils.verify(false,true,"Email Id is Needed to Create a Employee");
}

  
var ApproverGroup_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2) 
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
  
var EmploymentType_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 6).SWTObject("McPopupPickerWidget", "", 2);
if(EmploymentType!=""){
EmploymentType_1.Click();  
Delay(5000);
Sys.Process("Maconomy").Refresh(); 
WorkspaceUtils.DropDownList(EmploymentType,"Employment Type"); 

 }else{ 
  ValidationUtils.verify(false,true,"Employment Type is Needed to Create a Employee");
} 
  
  
var SalesEmployee_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
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


var EmployeeDepartment_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(EmployeeDepartment_Name!=""){
EmployeeDepartment_1.Click();
var ExlArray = getExcelData("ValidateDepartment",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation_Name_column2(EmployeeDepartment_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 3").OleValue.toString().trim(),EmployeeDepartment_Name,ExlArray,"Employee Department Number");
 }else{ 
  ValidationUtils.verify(false,true,"Employment Type is Needed to Create a Employee");
} 

    
var EmployeeCostCentre_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(EmployeeCostCentre_Name!=""){
EmployeeCostCentre_1.Click();
var ExlArray = getExcelData("Validate_BusinessUnit",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation_Name_column2(EmployeeCostCentre_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 5").OleValue.toString().trim(),EmployeeCostCentre_Name,ExlArray,"Business Unit Number");
}else{ 
  ValidationUtils.verify(false,true,"Employee Cost Centre is Needed to Create a Employee");
}

    
var Supervisor_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Supervisor!=""){
Supervisor_1.Click();
WorkspaceUtils.SearchByValue(Supervisor_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Supervisor,"Supervisor");

     }else{ 
  ValidationUtils.verify(false,true,"Supervisor Centre is Needed to Create a Employee");
} 

    
    
var AbsenceApprover_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(AbsenceApprover!=""){
AbsenceApprover_1.Click();
WorkspaceUtils.SearchByValue(AbsenceApprover_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),AbsenceApprover,"AbsenceApprover");

     }else{ 
  ValidationUtils.verify(false,true,"Absence Approver is Needed to Create a Employee");
}
    
  Delay(2000);
  var next = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
 Sys.HighlightObject(next);
 next.HoverMouse();
ReportUtils.logStep_Screenshot();
 next.Click();

}
}
TextUtils.writeLog("Details is entered in screen 1 and clicked NEXT"); 
}


function Employee_Information1(){ 
Delay(2000);
var Role_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
if(Role!=""){
 Role_1.Click();
var ExlArray = getExcelData("Validate_EmployeeCategories",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(Role_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee Category").OleValue.toString().trim(),Role,ExlArray,"Role");
// WorkspaceUtils.SearchByValue(Role_1,"Employee Category",Role,"Role");
     }else{ 
  ValidationUtils.verify(false,true,"Role is Needed to Create a Employee");
} 

  
var VacationCalendar_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McValuePickerWidget", "", 2);
if(VacationCalendar!=""){
VacationCalendar_1.Click();
 WorkspaceUtils.SearchByValue(VacationCalendar_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vacation Calendar").OleValue.toString().trim(),VacationCalendar,"Vacation Calendar");
     }else{ 
  ValidationUtils.verify(false,true,"Vacation Calendar is Needed to Create a Employee");
}

    
var WeekCalendarNo_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McValuePickerWidget", "", 2);
if(WeekCalendarNo!=""){
 WeekCalendarNo_1.Click();
    WorkspaceUtils.SearchByValue(WeekCalendarNo_1,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Week Calendar").OleValue.toString().trim(),WeekCalendarNo,"WeekCalendar No");

   }else{ 
  ValidationUtils.verify(false,true,"Week Calendar No is Needed to Create a Employee");
}

  

var CreateUser_1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "").OleValue.toString().trim());
if(CreateUser_1.getSelection()){ 
  CreateUser_1.Click();
   ValidationUtils.verify(true,true,"Create user is Unselected");
  }
else{ 
  ReportUtils.logStep("INFO","Create User Check Box is Already UnSelected");
}
    
    
Delay(2000); 
var submit = Aliases.Maconomy.Absence.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit").OleValue.toString().trim());

//Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Employee").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit").OleValue.toString().trim());
Sys.HighlightObject(submit);
submit.HoverMouse();
ReportUtils.logStep_Screenshot();
submit.Click();
Delay(6000); 
  
var label1 = Sys.Process("Maconomy").WaitSWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());

Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());

Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}

if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());

Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}

if(Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee"){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());

Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
TextUtils.writeLog("Details is entered in screen 2 and clicked Submit");
}





function Employess(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}  

var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1)
.SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Employees").OleValue.toString().trim());

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
TextUtils.writeLog("Created Employee is available in Maconomy screen");
  
  
if(flag){ 
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
empNumber  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 3).getText().OleValue.toString().trim();
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Maximize.Click();
var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 6,60000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
All_approver.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
TextUtils.writeLog("Finding approvers for Created Employee");
var info_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
info_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}




if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}



if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}


if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}


if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
ValidationUtils.verify(true,true,"Created Employee is Approved by :"+Project_manager)
TextUtils.writeLog("Created Employee is Approved by :"+Project_manager);
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(comapany+" - "+empNumber+" - Approver :"+Project_manager);
}
//}

if(Approve_Level.length==1){
var Employee_Vendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
Employee_Vendor.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
Sys.HighlightObject(table)
Sys.HighlightObject(closeFilter)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
  
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
TextUtils.writeLog("Created Employee Vendor is available in system"); 
  
  
if(flag){ 
closeFilter.Click();
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
All_approver.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Forward.Click();  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}



var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
All_approver.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Forward.Click();  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

}

}
}
}
}
}


function FinalApproveEmployess(comapany,empNumber,userNmae,apvLvl){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
//aqUtils.Delay(5000, Indicator.Text);
// Delay(5000) 

if(ImageRepository.ImageSet.Close_Filter.Exists()){ 
  
}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);  
}

//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//aqUtils.Delay(2000, Indicator.Text);
////Delay(5000)
//ImageRepository.ImageSet.Show_Filter.Click();
//aqUtils.Delay(2000, Indicator.Text);
////Delay(3000);
//} 

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(comapany);
//  firstCell.Keys("1707");
firstCell.Keys("[Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  job.setText(empNumber)
//job.setText(FullName + " "+STIME);
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
TextUtils.writeLog("Created Employee is available in Approver List");  
  
  
if(flag){ 
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
  
  Delay(8000); 
if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}




if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}



if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}


if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
Delay(8000); 
}


if((Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Employees - Employee Information") || (Sys.Process("Maconomy").SWTObject("Shell", "*", 1).WndCaption=="Empleados - Empleado Información")){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Employee Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
}
ValidationUtils.verify(true,true,"Created Employee is Approved by :"+userNmae)
TextUtils.writeLog("Created Employee is Approved by :"+userNmae); 
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(comapany+" - "+empNumber+" - Approver :"+userNmae);
}
Delay(4000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if(apvLvl==(ApproveInfo.length-1)){
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//if(ImageRepository.ImageSet.Forward.Exists()){
//ImageRepository.ImageSet.Forward.Click();// GL
//}

var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).WaitSWTObject("TabControl", "",1,60000);
var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).WaitSWTObject("TabControl", "", 6,60000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
All_approver.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
Sys.HighlightObject(table)
Sys.HighlightObject(closeFilter)

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

  
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
TextUtils.writeLog("Created Employee Vendor is available in system");  
  
  
if(flag){ 
closeFilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closeFilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot(); 
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
All_approver.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Forward.Click();  

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


var approve_Bar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
if(approve_Bar.isEnabled()){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot(); 
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 
  
  
  

var All_approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
All_approver.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Forward.Click();  

}

//}
}
}
}
}

















function ApproveEmployee(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
Sys.HighlightObject(table)
Sys.HighlightObject(closefilter)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
//  else{ 
//   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
//    if(UserN){ 
//      goToHR();
//      UserN = false;
//    }
//    temp = searchNumber(Eno);
//  }
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















function CreateEmployee(){ 
TextUtils.writeLog("Create Employee and Employee vendor Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);
  
}
excelName = EnvParams.path;
FullName,Gender,Country,comapany,DateEmployed,Position,Email,ApproverGroup,EmploymentType,SalesEmployee,EmployeeDepartment_Name,EmployeeCostCentre_Name,Supervisor,AbsenceApprover,Role,VacationCalendar,WeekCalendarNo,CreateUser,UserType,ValidityPeriodFrom,ValidityPeriodTo,AccessLevel,Language = "";
Approve_Level = [];
UserLevel = [];
ApproveInfo = [];
Emp_Vendor_Approve_Level = [];
Language = "";
empNumber = "";
Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
//Log.Message(EnvParams.Opco)
//Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
//Log.Message(Language)

sheetName = "EmployeeCreation";



STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create Employee test started::"+STIME);
TextUtils.writeLog("Create Employee test started::"+STIME);
try{

goToMenu(); 
goToEmployees();
getDetails();
//EmployeeScreen1_Address();
Employee_Information();
//EmployeeScreen2_Address();
Employee_Information1();
Employess();

WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();

for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApproveEmployess(temp[0],temp[1],temp[2],i);
}

}
catch(err){ 
  Log.Message(err);
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

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
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
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



