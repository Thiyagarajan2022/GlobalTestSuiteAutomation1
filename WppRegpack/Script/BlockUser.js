//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var userInfo = [];
Approve_Level = [];
y=0;
ApproveInfo = [];
level =0;
var UserPasswd = [];
var sheetName = "BlockUser";
//var sheetName1 = "Access Level";
//var loginpassword = "Credentials";
//var sscCredential = "SSC Credential";
var Credential = "userRoles";
var third_lvl_approver = false;
var login_satuts;
var STIME = "";
var LoginArr = [];
var HRData = [];
var LoginEmp = [];
//var lastArray = [];
//var Arrays = [];
var LoginArrays = [];
ExcelUtils.setExcelName(workBook, sheetName, true);
var nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,comapany,AccessLvel,EmployeeName;



function blockuser()
{
  Indicator.PushText("waiting for window to open");
//aqTestCase.Begin("Job Creation", "zfj://CH1-67");
excelName = EnvParams.getEnvironment();
workBook = Project.Path+excelName;
sheetName = "BlockUser";
Language = "";
Log.Message(workBook);
Log.Message(excelName);
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);

Arrays = [];
count = true;
checkmark = false;
STIME = "";
nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,EmployeeName,comapany="";

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
//Log.Message(EnvParams.Opco)
//Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
//Log.Message(Language)

    if(Language=="English"){
    STIME = WorkspaceUtils.StartTime();
    ReportUtils.logStep("INFO", "Block User started::"+STIME);

 
    getDetails();
//      goToMenu_user();
//    goToUsers();
//    searchforUser();
//    deleteUser();
//    closeAllWorkspaces();
//       goToMenu_employee();
//       Employeesearch();
//       closeAllWorkspaces();
         goToMenu_employee();
         searchForEmployeeVendor();
         
         
  closeAllWorkspaces();
 
  
}
}


function goToUsers(){ 
    Delay(3000)
    ReportUtils.logStep("INFO", "Enter User Details");
     var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
         All_User.HoverMouse();
ReportUtils.logStep_Screenshot("");
    
     All_User.Click();
    //  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
    //  employees.DblClickItem("|Employees");
      Delay(4000);
  //    address();
    }



function getDetails(){
  var sheetName = "BlockUser";
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(workBook);
//Log.Message(excelName);
nameValue = ExcelUtils.getRowDatas("Name",EnvParams.Opco)
//var Eml_split1 = nameValue.substring(0,nameValue.indexOf("@"));
//var Eml_split2 = nameValue.substring(nameValue.indexOf("@"));
//nameValue = Eml_split1 + " "+STIME+Eml_split2 
//nameValue = nameValue.replace(/[_: ]/g,""); 

if((nameValue==null)||(nameValue=="")){ 
ValidationUtils.verify(false,true,"Name is required to create USER");
}
//Log.Message("nameValue"+nameValue);
//employeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
//if((employeeNo==null)||(employeeNo=="")){ 
//ValidationUtils.verify(false,true,"EmployeeNo is required to create USER");
//}

employeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
//if((employeeNo==null)||(employeeNo=="")){ 
//employeeNo =readlog();
//Log.Message("jobNoNotepad= "+employeeNo);
//}
if((employeeNo==null)||(employeeNo=="")){ 
ValidationUtils.verify(false,true,"employeeNo is Needed to Create a USER");
}

//Log.Message(employeeNo);

userType = ExcelUtils.getRowDatas("User Type",EnvParams.Opco)
if((userType==null)||(userType=="")){ 
ValidationUtils.verify(false,true,"User Type is required to create USER");
}

//Log.Message("userType"+userType);

validFrom = ExcelUtils.getRowDatas("Valid From",EnvParams.Opco)
if((validFrom==null)||(validFrom=="")){ 
ValidationUtils.verify(false,true,"Valid From is required to create USER");
}
//Log.Message("validFrom"+validFrom);

validTo = ExcelUtils.getRowDatas("Valid To",EnvParams.Opco)
if((validTo==null)||(validTo=="")){ 
ValidationUtils.verify(false,true,"Valid To is required to create USER");
}

//Log.Message("validTo"+validTo);


AccessLevel = ExcelUtils.getRowDatas("Access Level",EnvParams.Opco)

if((AccessLevel==null)||(AccessLevel=="")){ 
//ValidationUtils.verify(false,true,"Acc is Needed to Create a USER");
}

EmployeeName = ExcelUtils.getRowDatas("Employee Name",EnvParams.Opco)

if((EmployeeName==null)||(EmployeeName=="")){ 
ValidationUtils.verify(false,true,"EmployeeName is Needed to Create a USER");
}

//Log.Message(AccessLevel);

//ExcelUtils.setExcelName(workBook, sheetName1, true);
companyNo = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

//Log.Message("companyNo"+companyNo);

ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}


      
      
      

    
    
    function searchforUser()
    {
      Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
 firstcell.Keys(nameValue);

     ReportUtils.logStep_Screenshot("");
//  firstcell.Keys("1307_AutomationUser 24September2019 16:20:30");
  firstcell.Keys("[Tab][Tab]");
  Delay(2000);
//  var Employee_no = 
//  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
// // Employee_no.Keys(userInfo[1]);  
//   Employee_no.Keys(employeeNo);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Delay(4000);  
    var flag=false;
    for(var v=0;v<table.getItemCount();v++){ 
  //    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userInfo[0]+" "+STIME))
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()=="Muthu 18:42:32"){ 
        flag=true;
        Log.Message("User is available");
        break;
      }else{ 
        table.Keys("[Down]");
      }
    }

//    var flag = table.getItemCount()>0;
    ValidationUtils.verify(flag,true,"User is available in system");
       ReportUtils.logStep_Screenshot("");
  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
   
    closefilter.HoverMouse();
    ReportUtils.logStep_Screenshot();
    closefilter.Click();
    Delay(3000);

    }
        Delay(6000);
   


    }
    
    function deleteUser()
    {
      Delay(3000);

var deleteuserButton = Aliases.ObjectGroup.deleteUserButton


deleteuserButton.Click();
Delay(3000);


var deleteUserPopupOk = Aliases.ObjectGroup2.DeleteUserPopUpOk

//deleteuserButton.Click();

Sys.HighlightObject(deleteUserPopupOk);

var deleteUserCancel = Aliases.ObjectGroup2.DeleteUserPopupCancel

deleteUserCancel.Click();

Delay(6000);
    }
    
    function goToMenu_user(){ 
  
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

//if(ImageRepository.ImageSet.User1.Exists()){
//  ImageRepository.ImageSet.User1.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.User3.Exists()){
//  ImageRepository.ImageSet.User3.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.User2.Exists()){
//  ImageRepository.ImageSet.User2.DblClick();// GL
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

      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
        var Client_Managt;
    //  Log.Message(childCC)
      for(var i=1;i<=childCC;i++){ 
      Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
      if(Client_Managt.isVisible()){ 
      Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
           Client_Managt.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Client_Managt.DblClickItem("|Users");
      
      }
      }
      Delay(3000);
      Log.Message("User Icon is Clicked");
      }
      
      
function goToMenu_employee(){ 
  
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
Delay(3000);

}


function goToEmployees(){ 

//  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
//  employees.DblClickItem("|Employees");

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
  
}

function Employeesearch(){ 
var All_Emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1)
.SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Employees");
All_Emp.Click();
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(companyNo);
//  firstCell.Keys("1707");
firstCell.Keys("[Tab][Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//  job.setText("GAIL C COUTINHO")
job.setText(EmployeeName );
Delay(6000);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(2).OleValue.toString().trim()==(EmployeeName)){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

if(flag){ 
closefilter.Click();
Delay(5000);
}
    
    
    
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Employee Created is available in system");
  
  Delay(5000);
  
  
  
  var terminationdate = Aliases.ObjectGroup.TerminationdateField;
  terminationdate.setText("11/15/2019")
  
  
  var blockedCheckbox =Aliases.ObjectGroup.blockedCheckBox;
  blockedCheckbox.Click();
    Delay(4000);
    
    savebutton = Aliases.ObjectGroup.SaveButtonBlockUser
    savebutton.Click();
    
      Delay(4000);
      
  var  UserResubmitPopupOk =   Aliases.UserResubmitPopupOk
  
   var  UserResubmitPopupcancel =   Aliases.ResubmitCancel
  
   UserResubmitPopupOk.Click();
 //UserResubmitPopupcancel.Click();
        Delay(4000);
   
//        var modifyChnagesYes = Aliases.ModifyChangesYes
//        
//           var modifyChnagesNo = Aliases.ModifyChangesNo
//      
//           //  modifyChnagesYes.Click();
// modifyChnagesNo.Click();    
   
Delay(5000);
 
//if(flag){ 
//closefilter.Click();
//Delay(5000);
//}

}
function test()
{
  
 var Blockeddropdown =Aliases.ObjectGroup.GlobalVendorInfo.BlockedDropdown;
Sys.HighlightObject(Blockeddropdown)
Blockeddropdown.Keys("")
Blockeddropdown.Click()
Blockeddropdown.Keys("Yes")
Blockeddropdown.Keys("[Up][Up]");
Blockeddropdown.Keys("[Enter]");


Delay(3000);
var StatusDropdown =Aliases.ObjectGroup.GlobalVendorInfo.StatusDropdown
Sys.HighlightObject(StatusDropdown)
StatusDropdown.Keys("")
StatusDropdown.Click()
Blockeddropdown.Keys("[Down][Down][Up]");
StatusDropdown.Keys("Inactive")
StatusDropdown.Keys("[Enter]");
}


function searchForEmployeeVendor()
{
     var  employeeVendorTab = Aliases.ObjectGroup.EmployeeVendorTab
      employeeVendorTab.Click();
//var Employee_Vendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
//Employee_Vendor.Click();
Delay(10000);

 Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//Aliases.ObjectGroup.McValuePickerWidget
//Aliases.ObjectGroup.VendorSearchusingcompanyName
//Aliases.ObjectGroup.firstCellEmployeeVendor;

// Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(companyNo);
Delay(4000);
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

//Aliases.ObjectGroup.VendorName

//Aliases.ObjectGroup.EmployeeNameField;

//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);

job.setText(EmployeeName);
//job.setText("GAIL C COUTINHO");
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var closeFilter = 
//
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
Delay(10000);
  
  
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(4).OleValue.toString().trim()==(EmployeeName)){ 
//      if(table.getItem(v).getText_2(4).OleValue.toString().trim()=="GAIL C COUTINHO"){
    flag=true;
    break;
  }else{ 
    table.Keys("[Down]");
  }
}
    
    
    
    
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Employee Vendor is available in system");
  

var CloseFilterEmployeeVendor = Aliases.ObjectGroup.CloseFilterSearchVendor;
CloseFilterEmployeeVendor.Click();

Delay(3000);

var GlobalVendorInfo =Aliases.ObjectGroup.GlobalVendorInfo;
GlobalVendorInfo.Click();
GlobalVendorInfo.MouseWheel(-300);

//closeFilter.Click();
Delay(3000);
var Blockeddropdown =Aliases.ObjectGroup.GlobalVendorInfo.BlockedDropdown;

Blockeddropdown.Keys("")
Blockeddropdown.Click()
Blockeddropdown.Keys("Yes")
Blockeddropdown.Keys("[Up][Up]");
//Blockeddropdown.Keys("[Enter]");

Delay(3000);
var StatusDropdown =Aliases.ObjectGroup.GlobalVendorInfo.StatusDropdown
StatusDropdown.Keys("")
StatusDropdown.Click()
Blockeddropdown.Keys("[Down][Down][Up]");
StatusDropdown.Keys("Inactive")
//StatusDropdown.Keys("[Enter]");
Delay(3000);

var EmployeeVendorSavebutton = Aliases.ObjectGroup.SaveButtonEmployeeVendor;
EmployeeVendorSavebutton.Click();
Delay(3000);

if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Employees - Global Vendor Information")    
{
  
Log.Message("Inside popup");
var button = Aliases.UserResubmitPopupOk
var buttoncancel =Aliases.ResubmitCancel
//Sys.Process("Maconomy").SWTObject("Shell", "Users - User Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Employees - Global Vendor Information").SWTObject("Label", "*").WndCaption;
      
 //Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
      Log.Message(label );
       //button.HoverMouse();
   //  ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(5000);
  } 

//var VendorOk = Aliases.UserResubmitPopupOk
//var vendorcancel = Aliases.ResubmitCancel
//vendorcancel.Click();

var companyVendorinfo = Aliases.ObjectGroup.CompanyVendorInfo;
companyVendorinfo.Click();
companyVendorinfo.MouseWheel(-300);

var companyvendorInfoBlocked =Aliases.ObjectGroup.CompanyVendorInfo.CompanyVendorinfoBlocked;

var companyVendorInfoStatus =NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.CompanyVendorInfo.Composite.McGroupWidget.CompanyVendorInfoStatus.Com;

Log.Message(companyvendorInfoBlocked.getText());
//Log.Message(companyVendorInfoStatus.getText());


}

 function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
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
Log.Message(compStatus);
return compStatus
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
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




 
