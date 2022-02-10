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
Language = "";
var UserPasswd = [];
var sheetName = "UserCreation";
//var sheetName1 = "JobCreation";
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

//var workBook = Project.Path+excelName;

ExcelUtils.setExcelName(workBook, sheetName, true);
var nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,comapany,AccessLvel;

//function mainscript(){ 
//employeeNo =  readlog();
//Log.Message(employeeNo);
//}

function getDetails(){
  var sheetName = "UserCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(workBook);
//Log.Message(excelName);
nameValue = ExcelUtils.getRowDatas("Name",EnvParams.Opco)
var Eml_split1 = nameValue.substring(0,nameValue.indexOf("@"));
var Eml_split2 = nameValue.substring(nameValue.indexOf("@"));
nameValue = Eml_split1 + " "+STIME+Eml_split2 
nameValue = nameValue.replace(/[_: ]/g,""); 

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

//Log.Message(AccessLevel);

//ExcelUtils.setExcelName(workBook, sheetName1, true);
companyNo = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

//Log.Message("companyNo"+companyNo);

validateCompany = getExcelData_Company("Validate_Company",EnvParams.Opco)
if((validateCompany==null)||(validateCompany=="")){ 
ValidationUtils.verify(false,true,"validateCompany is required to create USER");
}
//Log.Message("validateCompany"+validateCompany);




ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

}







function goToMenu(){ 
  
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
      
 function address(){
      Delay(4000);
      Sys.Process("Maconomy").Refresh();
      var name = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("McTextWidget", "", 1,60000).getText().OleValue.toString().trim();
      if(name!="Name")
        ValidationUtils.verify(false,true,"Name field is missing in Maconomy for Creation of User");
          else
        ValidationUtils.verify(true,true,"Name field is available in Maconomy for Creation of User");
      var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
      if(employee!="Employee No.")
        ValidationUtils.verify(false,true,"Employee field is missing in Macanomy for Creation of User");
        else
        ValidationUtils.verify(true,true,"Name field is available in Maconomy for Creation of User");
      var company = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
      if(company!="Company")
        ValidationUtils.verify(false,true,"Company field is missing in Macanomy for Creation of User");
        else
        ValidationUtils.verify(true,true,"Company field is available in Maconomy for Creation of User");
      var type = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
      if(type!="Type")
        ValidationUtils.verify(false,true,"Type field is missing in Macanomy for Creation of User");
        else
        ValidationUtils.verify(true,true,"Type field is available in Maconomy for Creation of User");
      var period = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
      if(period!="Period")
        ValidationUtils.verify(false,true,"Period field is missing in Macanomy for Creation of User");
        else
        ValidationUtils.verify(true,true,"Period field is available in Maconomy for Creation of User");
    }


function goToUsers(){ 
    Delay(3000)
    ReportUtils.logStep("INFO", "Enter User Details");
  //   var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
 //        All_User.HoverMouse();
   //   All_User.Click();
ReportUtils.logStep_Screenshot("");
    

    //  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
    //  employees.DblClickItem("|Employees");
      var Add_Visible0 = true;
      var New_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 3,60000);
      while(Add_Visible0){
      if(New_User.isEnabled()){
        Delay(2000);
                New_User.HoverMouse();
ReportUtils.logStep_Screenshot("");
      New_User.Click();
      Add_Visible0 = false;
      }
      }
      Delay(4000);
  //    address();
    }


function UserInformation(){ 
//userInfo =  excel(sheetName);

//Log.Message("User Creation is Started");
//
//  Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("McTextWidget", "", 2,60000);
//  Delay(1000);
//  var name = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//  if(userInfo[0]!=""){
//  name.Keys("^a[BS]")
//  name.setText(userInfo[0]+" "+STIME);
//  }else{ 
//    ValidationUtils.verify(false,true,"User Name is Needed to Create a User");
//  }


///----Entering  Name

  if(nameValue!=""){
  var Name_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//    Name_1.HoverMouse();
//ReportUtils.logStep_Screenshot("");
  Name_1.Click();
   Delay(2000);
  Name_1.setText(nameValue);
  ValidationUtils.verify(true,true,"Name is entered in Maconomy");
//  ReportUtils.logStep_Screenshot("");
  }
  else{ 
    ValidationUtils.verify(false,true,"Name is Needed to Create a Employee");
  }

  
////---Entering Employee Number
  var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2) 
  if(employeeNo!=""){
//        employee.HoverMouse();
//ReportUtils.logStep_Screenshot("");
    employee.Click();
    WorkspaceUtils.SearchByValue(employee,"Employee",employeeNo,"Employee Number"); 
//    ReportUtils.logStep_Screenshot("");
  } 
  else{ 
  ValidationUtils.verify(false,true,"Employee Number is Needed to Create a User");
  }
   

///------Entering Company Number----

//Log.Message(companyNo);
//Log.Message(validateCompany);

   var company = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,6000);
    if(companyNo!=""){
//         company.HoverMouse();
//ReportUtils.logStep_Screenshot("");
    company.Click();
    var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
    WorkspaceUtils.config_with_Maconomy_Validation(company,"Company",companyNo,validateCompany,"Company No.");
//    ReportUtils.logStep_Screenshot("");

  //  WorkspaceUtils.config_with_Maconomy_Validation(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),companyNo,ExlArray,"Company Number");
  }else  { 
      ValidationUtils.verify(false,true,"Company Name is Needed to Create a User");
    }
    
    
//    //----------Entering Company Number-------------
//  var companyName = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,60000);
//if(comapany!=""){
//companyName.Click();
//var ExlArray = getExcelData("Validate_Company",EnvParams.Opco)
//WorkspaceUtils.config_with_Maconomy_Validation(companyName,"Company",comapany,ExlArray,"Company Number");
//}else{ 
//  ValidationUtils.verify(false,true,"Country is Needed to Create Job");
//}
    


///------Entering User Type---- 
   var user_type = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
   if(userType!=""){
//     user_type.HoverMouse();
//     ReportUtils.logStep_Screenshot("");
   user_type.Click();
   WorkspaceUtils.SearchByValue(user_type,"User Type",userType,"User Type");
//    ReportUtils.logStep_Screenshot("");
   }
   else{ 
    ValidationUtils.verify(false,true,"User Number is Needed to Create a User");
   }
       
  var periodfrom = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 2)
  if(validFrom!=""){
    WorkspaceUtils.CalenderDateSelection(periodfrom,validFrom)
//     ReportUtils.logStep_Screenshot("");
  }

  var periodto = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 4)
  if(validTo!=""){
  WorkspaceUtils.CalenderDateSelection(periodto,validTo)

//     ReportUtils.logStep_Screenshot("");
  }  

  var template = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "").SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  if((template.getSelection()) && ((userInfo[6]=="No")||(userInfo[6]=="no"))) { 
    template.Click();
      Log.Message("Templete is UnChecked")
    }
  if((!template.getSelection()) && ((userInfo[6]=="Yes")||(userInfo[6]=="Yes"))) { 
    template.Click();
      Log.Message("Templete is Checked")
    }
  
  Delay(2000);
  var create = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
  Sys.HighlightObject(create);
   create.HoverMouse();
     ReportUtils.logStep_Screenshot("");
  create.Click();
  ReportUtils.logStep("INFO",+nameValue+" : is Created");
  Delay(8000);  
//  var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
//  Sys.HighlightObject(cancel);

var index = 0;
var MacCount = Sys.Process("Maconomy").ChildCount;
for(var mc =0;mc<MacCount;mc++){ 
//Log.Message(Sys.Process("Maconomy").Child(mc).Name)
if((Sys.Process("Maconomy").Child(mc).Name.indexOf("SWTObject(")!=-1)&&(Sys.Process("Maconomy").Child(mc).JavaClassName=="Shell")){ 
  if(Sys.Process("Maconomy").Child(mc).WndCaption=="Users - Users"){ 
 //   Log.Checkpoint(Sys.Process("Maconomy").Child(mc).FullName)
    index++;
  }
}

}


//for(var mc =1;mc<=index;mc++){ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",mc).isVisible()){
//var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",mc).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",mc).SWTObject("Label", "*");
//      Log.Message(label.getText());
//      button.Click();
//      Delay(5000);
//      }
//      }

if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
      Log.Message(label.getText());
       button.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(5000);
  }
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){    
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
      Log.Message(label.getText());
       button.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(5000);
 }
 
 if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){     
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
      Log.Message(label.getText());
       button.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(5000);
 }
 
//  if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){     
//var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
//      Log.Message(label.getText());
//       button.HoverMouse();
//     ReportUtils.logStep_Screenshot("");
//      button.Click();
//      Delay(5000);
// }

      
return true; 

    if(index==1){ 
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users").SWTObject("Text", "");
      Log.Message(label.getText());
      ImageRepository.ImageSet.OK_Button.Click();
      Delay(3000);
    
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
      Sys.HighlightObject(cancel);
      Log.Message("Create button is Invisible");
       cancel.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      cancel.Click();      
           
      ValidationUtils.verify(false,true,"User is Not Created");
      return false;
    }
}
   
  
function user_Approval(){ 
  Delay(4000);
  
   var blockedUserButton =  Aliases.ObjectGroup.BlockedUserRadioButton;
   blockedUserButton.Click();
     Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
 firstcell.Keys(nameValue);

 
 var blockedUserButton =  Aliases.ObjectGroup.BlockedUserRadioButton;

  firstcell.Keys("[Tab][Tab]");
  Delay(2000);

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Delay(4000);  
    var flag=false;
    for(var v=0;v<table.getItemCount();v++){ 
  //    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userInfo[0]+" "+STIME))
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()=="Muthu 18:42:32"){ 
        flag=true;
        Log.Message("User is created");
        break;
      }else{ 
        table.Keys("[Down]");
      }
    }

//    var flag = table.getItemCount()>0;
    ValidationUtils.verify(flag,true,"User Created is available in system");
       ReportUtils.logStep_Screenshot("");
  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
   
    closefilter.HoverMouse();
    ReportUtils.logStep_Screenshot();
    closefilter.Click();
    Delay(3000);

    }
    
    if((AccessLevel!="")&&(AccessLevel!=null)){ 
      
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    var Access_Level = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
    Sys.HighlightObject(Access_Level);
       Access_Level.HoverMouse();
    ReportUtils.logStep_Screenshot();
    Access_Level.Click();
    Delay(3000);
    
  var add = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
     add.HoverMouse();
    ReportUtils.logStep_Screenshot();
  add.Click();
  Delay(5000);
  
  var cell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//  SearchByValues_Col_1(ObjectAddrs,popupName,value,fieldName)

   cell.Click();
   Delay(3000);
   WorkspaceUtils.AccessLevel_Add(cell,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Access Level").OleValue.toString().trim(),AccessLevel,"AccessLevel");

  var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
 
      save.HoverMouse();
    ReportUtils.logStep_Screenshot();
  save.Click();
}
Delay(3000);
    

   var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
   Sys.HighlightObject(submit);
     submit.HoverMouse();
    ReportUtils.logStep_Screenshot();
    submit.Click();
    Delay(5000);
    var approval = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
        approval.HoverMouse();
    ReportUtils.logStep_Screenshot();
    approval.Click();
    Delay(3000);
    ImageRepository.ImageSet.Maximize.Click();
    Delay(2000);
    var All_approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
         All_approve.HoverMouse();
    ReportUtils.logStep_Screenshot();
   
    All_approve.Click();
    
//   var apprve_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
   var apprve_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);;
   var y=0;
  for(var x=0;x<apprve_table.getItemCount();x++){ 
      third_lvl_approver = true;
      var approvers="";
       if(apprve_table.getItem(x).getText_2(8)!="Approved"){
       approvers = apprve_table.getItem(x).getText_2(3).OleValue.toString().trim()+"*"+apprve_table.getItem(x).getText_2(4).OleValue.toString().trim();
       }
       
  
                       
    }
    
       getApproverDetails();
        ReportUtils.logStep_Screenshot();
       
//    if(ImageRepository.ImageSet.Forward.Exists()){
//ImageRepository.ImageSet.Forward.Click();// GL
//}


var userInfoTab =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
       userInfoTab.HoverMouse();
    ReportUtils.logStep_Screenshot();
userInfoTab.Click();

var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
  
level = 1;
var Approve = "";
Approve =
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
Approve.HoverMouse();

    ReportUtils.logStep_Screenshot();
/////-----------------------------------------
Approve.Click();

ValidationUtils.verify(true,true,"Created User is Approved by" +Project_manager)
 ReportUtils.logStep("INFO", "Created User is Approved by" +Project_manager);
 
 
if(Approve_Level.length==1){
  Delay(5000);
  var Ok = 
  Sys.Process("Maconomy").SWTObject("Shell", "Users", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  //Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
 // Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  //'
    Ok.HoverMouse();
    ReportUtils.logStep_Screenshot();
  Ok.Click();
   Delay(3000);
   
   
  if(Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
//Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
    var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
 
 expDate.HoverMouse();
    ReportUtils.logStep_Screenshot();
  expDate.Click();
  
  
  }
  }
// 
//  
  
//--------------------------------------------------- 
//  level++;
  
    }
    }
    
function getApproverDetails()
{
 // if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).isVisible()){
    var Approval_table = 
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
  //  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    linestatus = true;
 //   }

    Sys.HighlightObject(Approval_table)
    for(var z=0;z<Approval_table.getItemCount();z++){ 
       approvers="";
       approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +z+ ": " +approvers);
//       Approve_Level[y] = comapany+"*"+nameValue+"*"+approvers;
       Approve_Level[y] = approvers;
       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
       y++;
    }

}



function goToHR(){ 
  Delay(3000);
    closeAllWorkspaces();


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

  

function gotoApprove(){
Delay(3000)
var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click(); 
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  firstcell.Keys(userInfo[0]+" "+STIME);
//  firstcell.Keys("Muthu");
  firstcell.Keys("[Tab][Tab]");
  Delay(2000);
  var Employee_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
//  Employee_no.Keys("170710134");
  Employee_no.Keys(userInfo[1]);
  Delay(4000);
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userInfo[0]+" "+STIME)){ 
      flag=true;
      break;
    }else{ 
      table.Keys("[Down]");
    }
  }  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    closefilter.Click();
    Delay(3000);
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
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
    DDT.CloseDriver(xlDriver.Name);
return Arrays;
}


function Access_lvl_excel(){ 
 var Arrays = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length-1;idx++){  
       if((xlDriver.Value(colsList[idx])!=null)&& (xlDriver.Value(colsList[idx]).toString().trim()=="Access Level")){
          Arrays[id]=xlDriver.Value(colsList[idx+1]).toString().trim();
          id++;
      }
         }
     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}



function RestMaconomy(){ 
for(var i=0;i<UserPasswd.length;i++){

var temp = UserPasswd[i];
temp_user = [];
temp_user = temp.split("*");
var uname = temp_user[2]; 
//Log.Message(uname)
var pwd = temp_user[3];
//Log.Message(pwd)
Rests(uname,pwd);
    
goToMenu();
gotoApprove();
ValidationUtils.verify(true,true,"User has Approved by "+uname);
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



function CredentialLogin(){ 

for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  Log.Message(Approve_Level[i]);
  var Cred = Approve_Level[i].split("*");
  for(var j=0;j<2;j++){
    Log.Message(Cred[j])
Log.Message(j)
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
//     var workBook = "H:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\China\\DS_SYSTEST_EN.xls";
     workBook = Project.Path+excelName;
     Log.Message(workBook)
     var sheetName = "Agency Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 
//    var workBook = "H:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\China\\DS_SYSTEST_EN.xls";
    workBook = Project.Path+excelName;
     Log.Message(workBook)
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
  Log.Message("Logins :"+temp);
}
WorkspaceUtils.closeAllWorkspaces();

}


function CreateUser(){ 
  

  Indicator.PushText("waiting for window to open");
//aqTestCase.Begin("Job Creation", "zfj://CH1-67");
excelName = EnvParams.getEnvironment();
workBook = Project.Path+excelName;
sheetName = "UserCreation";
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
nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,comapany="";

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
    ReportUtils.logStep("INFO", "Job Creation started::"+STIME);
   goToMenu();
    getDetails();
    goToUsers(); 
//  userInfo =  excel(sheetName);
  var status = UserInformation();
//var status = true;
  if(status){
  user_Approval();

//Approve_Level[0] = " Muthu 22:06:31*170710134*SACHINDRA P KARKERA (170710011)*";
  closeAllWorkspaces();
//  Approver level :0: 1307 Senior Finance (13079510)*1307 Management (13079507)	15:51:02	Normal			2.35
 // Approve_Level[1] = "1307"+"*"+"1307_AutomationUser 18October2019 12:58:29"+"*"+"1307 Senior Finance (13079510)*1307 Management (13079507)"; 
 
 //Approve_Level[0] = "1307"+"*"+"1307_AutomationUser 18October2019 12:58:29"+"*"+"1307 Senior Finance (13079510)*1307 Management (13079507)"; 
// Approve_Level[1] = "1307"+"*"+"1307_AutomationUser 18October2019 12:58:29"+"*"+"1307 Senior Finance (13079510)*1307 Management (13079507)"; 
  CredentialLogin();

for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
//Delay(20000);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
//Delay(5000);

todo(temp[3]);
aprvUser(temp[0],temp[1],temp[2],temp[3],i);


}
  

  }
  WorkspaceUtils.closeAllWorkspaces();
}
}


function aprvUser(ComId,EmpNo,userNmae,lvls,apvLvl){ 
aqUtils.Delay(5000, Indicator.Text);
// Delay(5000) 
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
//Delay(5000)
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);
//Delay(3000);
} 

  
    aqUtils.Delay(8000, Indicator.Text);
//    Delay(2000);

Delay(3000)
//var All_User =
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
// All_User.HoverMouse();
    ReportUtils.logStep_Screenshot();

//All_User.Click(); 
  Delay(4000);
  var firstcell = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")

   firstcell.Keys(nameValue);
  
//  firstcell.Keys("Muthu");
  firstcell.Keys("[Tab][Tab]");
  Delay(2000);
  var Employee_no = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2)

  Employee_no.Keys(employeeNo);

    ReportUtils.logStep_Screenshot();
  Delay(4000);
  
  var table = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)

  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 

        if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue)){  
      flag=true;
      break;
    }else{ 
      table.Keys("[Down]");
    }
  } 
  
     ValidationUtils.verify(flag,true,"User is listed for Approval"); 
  if(flag){ 
   var closefilter = 
   Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

    closefilter.Click();

    ReportUtils.logStep_Screenshot();
    Delay(5000);
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var approve =  Aliases.ObjectGroup.ApproveEmailUser;
  
// Aliases.ObjectGroup.ApproveAndEmailUser
  
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8)

  Sys.HighlightObject(approve);

 if(approve.isEnabled()){ 
   approve.HoverMouse();
    ReportUtils.logStep_Screenshot();
  approve.Click();
ValidationUtils.verify(true,true,"User is Approved by "+userNmae)
 ReportUtils.logStep("INFO", "user is Approved by "+userNmae);
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
//  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
}
  Delay(5000);
  
  if(apvLvl==(ApproveInfo.length-1)){
if(lvls==0){
  


  var Ok = 
  Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")

  Ok.HoverMouse();
    ReportUtils.logStep_Screenshot();
  Ok.Click();
   Delay(2000);
  if(Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
    

    var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
   expDate.HoverMouse();
  ReportUtils.logStep_Screenshot();
  expDate.Click();
  }
  
  }
 if(lvls==1){
   
 var Ok = 
 // Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information (Substitute)", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
 // Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  //
    Ok.HoverMouse();
    ReportUtils.logStep_Screenshot();
  Ok.Click();
   Delay(2000);
  if(Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information (Substitute)").SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {

    var expDate =
      Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information (Substitute)").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
   //  Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    expDate.HoverMouse();
  ReportUtils.logStep_Screenshot();
  expDate.Click();
  }
   
   }
  
  }
  
  }
  

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
if(lvl==0)
Client_Managt.DblClickItem("|Approve User Information (*)");

if(lvl==1)
Client_Managt.DblClickItem("|Approve User Information (Substitute) (*)");


break;
}
}
}



}

 function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
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

function readlog(){

sheetName = "JobCreation";

ExcelUtils.setExcelName(workBook, sheetName, true);

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

sheetName = "UserCreation";

ExcelUtils.setExcelName(workBook, sheetName, true);


}