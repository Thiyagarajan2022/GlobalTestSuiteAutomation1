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
var sheetName = "UserCreation";
var sheetName1 = "JobCreation";
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
var nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,comapany;

//function mainscript(){ 
//employeeNo =  readlog();
//Log.Message(employeeNo);
//}

function getDetails(){
  var sheetName = "UserCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
Log.Message(workBook);
Log.Message(excelName);
nameValue = ExcelUtils.getRowDatas("Name",EnvParams.Opco)
if((nameValue==null)||(nameValue=="")){ 
ValidationUtils.verify(false,true,"Name is required to create USER");
}
Log.Message("nameValue"+nameValue);
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

Log.Message(employeeNo);

userType = ExcelUtils.getRowDatas("User Type",EnvParams.Opco)
if((userType==null)||(userType=="")){ 
ValidationUtils.verify(false,true,"User Type is required to create USER");
}

Log.Message("userType"+userType);

validFrom = ExcelUtils.getRowDatas("Valid From",EnvParams.Opco)
if((validFrom==null)||(validFrom=="")){ 
ValidationUtils.verify(false,true,"Valid From is required to create USER");
}
Log.Message("validFrom"+validFrom);

validTo = ExcelUtils.getRowDatas("Valid To",EnvParams.Opco)
if((validTo==null)||(validTo=="")){ 
ValidationUtils.verify(false,true,"Valid To is required to create USER");
}

Log.Message("validTo"+validTo);

ExcelUtils.setExcelName(workBook, sheetName1, true);
companyNo = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

Log.Message("companyNo"+companyNo);

validateCompany = getExcelData_Company("Validate_Company",EnvParams.Opco)
if((validateCompany==null)||(validateCompany=="")){ 
ValidationUtils.verify(false,true,"validateCompany is required to create USER");
}
Log.Message("validateCompany"+validateCompany);


ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}



function getDetailsAdditional(){
  
//ExcelUtils.setExcelName(workBook, sheetName1, true);
//companyNo = ExcelUtils.getRowDatas("company",EnvParams.Opco)
//if((companyNo==null)||(companyNo=="")){ 
//ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
//}
//
//Log.Message("companyNo"+companyNo);
//
//validateCompany = ExcelUtils.getRowDatas("Validate_Company",EnvParams.Opco)
//if((validateCompany==null)||(validateCompany=="")){ 
//ValidationUtils.verify(false,true,"validateCompany is required to create USER");
//}
//Log.Message("validateCompany"+validateCompany);

//ExcelUtils.setExcelName(workBook, sheetName, true);

}



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

//if(ImageRepository.ImageSet.User1.Exists()){
//  ImageRepository.ImageSet.User1.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.User3.Exists()){
//  ImageRepository.ImageSet.User3.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.User2.Exists()){
//  ImageRepository.ImageSet.User2.DblClick();// GL
//}
      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
        var Client_Managt;
      Log.Message(childCC)
      for(var i=1;i<=childCC;i++){ 
      Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
      if(Client_Managt.isVisible()){ 
      Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
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
     var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
     All_User.Click();
    //  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
    //  employees.DblClickItem("|Employees");
      var Add_Visible0 = true;
      var New_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 3,60000);
      while(Add_Visible0){
      if(New_User.isEnabled()){
        Delay(2000);
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
  Name_1.Click();
   Delay(2000);
  Name_1.setText(nameValue + " "+STIME);
  ValidationUtils.verify(true,true,"Name is entered in Maconomy");
  }
  else{ 
    ValidationUtils.verify(false,true,"Name is Needed to Create a Employee");
  }

  
////---Entering Employee Number
  var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2) 
  if(employeeNo!=""){
    employee.Click();
    WorkspaceUtils.SearchByValue(employee,"Employee",employeeNo,"Employee Number"); 
  } 
  else{ 
  ValidationUtils.verify(false,true,"Employee Number is Needed to Create a User");
  }
   

///------Entering Company Number----

Log.Message(companyNo);
Log.Message(validateCompany);

   var company = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,6000);
    if(companyNo!=""){
    company.Click();
    var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
    WorkspaceUtils.config_with_Maconomy_Validation(company,"Company",companyNo,validateCompany,"Company No.");
    

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
   user_type.Click();
   WorkspaceUtils.SearchByValue(user_type,"User Type",userType,"User Type");
   }
   else{ 
    ValidationUtils.verify(false,true,"User Number is Needed to Create a User");
   }
       
  var periodfrom = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 2)
  if(validFrom!=""){
    WorkspaceUtils.CalenderDateSelection(periodfrom,validFrom)
  }

  var periodto = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 4)
  if(validTo!=""){
  WorkspaceUtils.CalenderDateSelection(periodto,validTo)
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
  create.Click();
  //ReportUtils.logStep("INFO", nameValue+" "+STIME +" : is Created");
  Delay(8000);  
//  var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
//  Sys.HighlightObject(cancel);

var index = 0;
var MacCount = Sys.Process("Maconomy").ChildCount;
for(var mc =0;mc<MacCount;mc++){ 
Log.Message(Sys.Process("Maconomy").Child(mc).Name)
if((Sys.Process("Maconomy").Child(mc).Name.indexOf("SWTObject(")!=-1)&&(Sys.Process("Maconomy").Child(mc).JavaClassName=="Shell")){ 
  if(Sys.Process("Maconomy").Child(mc).WndCaption=="Users - Users"){ 
    Log.Checkpoint(Sys.Process("Maconomy").Child(mc).FullName)
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
      button.Click();
      Delay(5000);
  }
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){    
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
      Log.Message(label.getText());
      button.Click();
      Delay(5000);
 }
 
 if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){     
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
      Log.Message(label.getText());
      button.Click();
      Delay(5000);
 }
 
 if(Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",2).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible()){     
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",2).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",2).SWTObject("Label", "*");
      Log.Message(label.getText());
      button.Click();
      Delay(5000);
 }     
      
return true; 
//
//  var pop = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users");
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Users - Users"){
//  if(pop.Child(1).JavaFullClassName.indexOf('Label')!=-1){
//    var index = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",).Index
//      var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Label", "*");
//      Log.Message(label.getText());
//      button.Click();
////      ImageRepository.ImageSet.OK_Button.Click();
//}
//
//      Delay(4000);
//      if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Users - Users")
//  if(pop.Child(1).JavaFullClassName.indexOf('Label')!=-1){
//    var index = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",).Index
//      var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Label", "*");
//      Log.Message(label.getText());
//      button.Click();
////      ImageRepository.ImageSet.Ok.Click();
//}
//
//      Delay(4000);
//      if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Users - Users")
//  if(pop.Child(1).JavaFullClassName.indexOf('Label')!=-1){
//    var index = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",).Index
//      var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Label", "*");
//      Log.Message(label.getText());
//      button.Click();
////      ImageRepository.ImageSet.Ok.Click();
//}
//
//      Delay(4000);
//      if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Users - Users")
//  if(pop.Child(1).JavaFullClassName.indexOf('Label')!=-1){
//    var index = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",).Index
//      var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Label", "*");
//      Log.Message(label.getText());
//      button.Click();
////      ImageRepository.ImageSet.Ok.Click();
//}
////      Delay(2000);
//////      var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",3).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//////      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",3).SWTObject("Label", "*");
////
////      Log.Message(label.getText());
////      //button.Click();
////      ImageRepository.ImageSet.Ok.Click();
////
////        Delay(2000);
////      //Sys.Process("Maconomy").SWTObject("Shell", "Users - Users", 4).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
////      if(ImageRepository.ImageSet.Ok.Exists()){
////      var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
////      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",index).SWTObject("Label", "*");
////      Log.Message(label.getText());
////      //button.Click();
////      ImageRepository.ImageSet.Ok.Click();
////        Delay(2000);
////      } 
//      Log.Message("User is Created");
//      return true; 
//    }
    if(index==1){ 
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users").SWTObject("Text", "");
      Log.Message(label.getText());
      ImageRepository.ImageSet.OK_Button.Click();
      Delay(3000);
    
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
      Sys.HighlightObject(cancel);
      Log.Message("Create button is Invisible");
      cancel.Click();      
           
      ValidationUtils.verify(false,true,"User is Not Created");
      return false;
    }
}
   
  
function user_Approval(){ 
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
 firstcell.Keys(nameValue+" "+STIME);
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
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue+" "+STIME)){ 
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
  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    closefilter.Click();
    Delay(3000);
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    var Access_Level = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
    Sys.HighlightObject(Access_Level);
    Access_Level.Click();
    Delay(3000);
    
    var Acc_lvl = Access_lvl_excel();
    for(var z=0;z<Acc_lvl.length;z++){
    var Add_Click = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(Add_Click);
    Add_Click.Click();
    Delay(3000);
//    if(z==0){
    var Acc_Code = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    }else{
//    var Acc_Code = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    }
    Sys.HighlightObject(Acc_Code);
    Acc_Code.Click();
//    Log.Message(Acc_lvl[z]);
    
    
  /*    
    
    Acc_Code.Keys("^a[BS]");
    Acc_Code.Keys(Acc_lvl[z]);
    Delay(4000);
    
    
    
  
    var op = Sys.Process("Maconomy").SWTObject("Shell", "");
    var sea = "Sys.Process(Maconomy).SWTObject(Shell, ).SWTObject(McValuePickerFooter, , 1)";
    var sea1 = "Sys.Process(Maconomy).SWTObject(Shell, ).SWTObject(McValuePickerFooter, , 1).SWTObject(CLabel, , 1)";
    var status = "fail";
    for(var i=0;i<op.ChildCount;i++){ 
    var test = op.Child(i).FullName.toString().trim();

    test = test.replace(/"/g,"");
    if(test==sea){ 
    var tt= op.Child(i).FullName;

    for(var j=0;j<op.Child(i).ChildCount;j++){ 
      var test1 = op.Child(i).Child(j).FullName;
       Log.Message(test1);
        test1 = test1.replace(/"/g,"");
      if(test1==sea1){ 
    status = "pass";
    }
}
}
}
    if(status == "pass"){
    var search = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("McValuePickerFooter", "", 1).SWTObject("CLabel", "", 1)
    Sys.HighlightObject(search);
    search.Click();
//    var table =
    Delay(2000);
*/

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "JobCreation";
    Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//    Log.Message("z :"+z +"  Acc_lvl[z]"+Acc_lvl[z]);
    code.setText(Acc_lvl[z]);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
//    Log.Message(table.getItemCount());
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Acc_lvl[z]){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          Delay(2000);
//       if(z==0){
       var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//       }else{
//       var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//       }
       save.Click();
       Delay(5000);
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          Acc_Code.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Option").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
      Acc_Code.setText("");
    } 
/*
    }
     if(status == "fail"){
      Acc_Code.setText("");
      Log.Warning("This value is not available in Maconomy for Access Level");
    }
  */  
    }

//   var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
   var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
   Sys.HighlightObject(submit);
    submit.Click();
    Delay(5000);
    var approval = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
    approval.Click();
    Delay(3000);
    ImageRepository.ImageSet.Maximize.Click();
    Delay(2000);
    var All_approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    All_approve.Click();
    
//   var apprve_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
   var apprve_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);;
   var y=0;
  for(var x=0;x<apprve_table.getItemCount();x++){ 
      third_lvl_approver = true;
      var approvers="";
       if(apprve_table.getItem(x).getText_2(8)!="Approved"){
       approvers = apprve_table.getItem(x).getText_2(3).OleValue.toString().trim()+"*"+apprve_table.getItem(x).getText_2(4).OleValue.toString().trim();
//       Log.Message("Approver level :" +z+ ": " +approvers);
      Log.Message("User Approver level : " +x+ " Approver :" +approvers);
//       if(userInfo[1]!=""){
//       Approve_Level[y] = userInfo[0]+" "+STIME+"*"+userInfo[1]+"*"+approvers;
//       }else{ 
//       Approve_Level[y] = userInfo[0]+" "+STIME+"**"+approvers;  
//       }
//        if(nameValue!=""){
//          Approve_Level[y] = nameValue +" "+STIME+"*"+employeeNo+"*"+approvers;
//        } 
//        else{
//          Approve_Level[y] = nameValue +" "+STIME+"*"+approvers;
//        } 
      // Log.Message(Approve_Level[y])
   
       y++;
       }
    }
   
  }
  
     getApproverDetails();
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
       Approve_Level[y] = comapany+"*"+employeeNo+"*"+approvers;
       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
       y++;
    }

}

function Login_Match(){ 
login_satuts = true;
Delay(3000);
UserPasswd = [];

goToHR();

Credentiallogin();

var z =0;
for(var i=0;i<Approve_Level.length;i++){ 
//if((Approve_Level[i].indexOf("OpCo")!=-1) && (userInfo[2]!="")){
//Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,userInfo[2]);
//}
if((Approve_Level[i].indexOf("OpCo")!=-1) && (userInfo[1]!="")){
Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,userInfo[1]);
}
// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
  if(Approve_Level[i].indexOf("SSC - Biller")==-1){
  Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
  }

var tempLevel = Approve_Level[i].split("*");
ifGotIT = true;
for(var j=2;j<tempLevel.length;j++){ 

if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
//Log.Message("temp :"+temp)
var tempSplit = temp.split("*");

  for(var k=0;k<LoginEmp.length;k++){
    var A_temp = LoginEmp[k].split("*");
//    Log.Message("tempSplit[0] :"+tempSplit[0]);
//    Log.Message("A_temp[0] :"+A_temp[0]);
//    Log.Message("tempSplit[1] :"+tempSplit[1]);
//    Log.Message("A_temp[1] :"+A_temp[1]);
   if(tempSplit[0]==A_temp[0]){ 
      UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
     Log.Message(UserPasswd[z]);
     z++;
     ifGotIT = false;
     break;     
   }else{ 
   if(tempSplit[1]==A_temp[2]){ 
      UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
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
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
     
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
//Log.Message("temp :"+temp)
var tempSplit = temp.split("*");

  for(var k=0;k<HRData.length;k++){
    var A_temp = HRData[k].split("*");
//    Log.Message("tempSplit[0] :"+tempSplit[0]);
//    Log.Message("A_temp[0] :"+A_temp[0]);
//    Log.Message("tempSplit[1] :"+tempSplit[1]);
//    Log.Message("A_temp[1] :"+A_temp[1]);
   if(tempSplit[1]==A_temp[1]){ 
     UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
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

return login_satuts;
}


function login() {
    var xlDriver = DDT.ExcelDriver(workBook, sscCredential, true);
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
     LoginArrays[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}


//function Credentiallogin() {
//    var xlDriver = DDT.ExcelDriver(Project.Path+ExcelFileName, Credential, true);
//var id =0;
//var colsList = [];
//
//   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
//     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
//   }
//     while (!DDT.CurrentDriver.EOF()) {
//     var temp ="";
//      for(var idx=0;idx<colsList.length;idx++){  
//       if(xlDriver.Value(colsList[idx])!=null){
//      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
//      }
//      else{ 
//        temp = temp+"*";
//      }
//      }
////      Log.Message(temp)
//     LoginEmp[id]=temp;
//     id++;     
//     xlDriver.Next();
//     }
//     DDT.CloseDriver(xlDriver.Name);
//}

function logins() {
    var xlDriver = DDT.ExcelDriver(workBook, loginpassword, true);
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
     LoginArr[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
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
    Log.Message(Cred[j])
Log.Message(j)
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf("1307"+" ")!=-1)))
  { 
//     var workBook = "H:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\China\\DS_SYSTEST_EN.xls";
     var sheetName = "Agency Users";
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


function CreateUser(){ 
  Indicator.PushText("waiing for window to open");
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


    Language = EnvParams.Language;
    if((Language==null)||(Language=="")){
      ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
    }
    Log.Message(Language)
    if(Language=="English"){
    STIME = WorkspaceUtils.StartTime();
    ReportUtils.logStep("INFO", "Job Creation started::"+STIME);

   goToMenu();
  // mainscript();
    getDetails();
    getDetailsAdditional();
    goToUsers();
    
//  userInfo =  excel(sheetName);
  var status = UserInformation();
//var status = true;
  if(status){
  user_Approval();
//Approve_Level[0] = " Muthu 22:06:31*170710134*SACHINDRA P KARKERA (170710011)*";
  closeAllWorkspaces();
  CredentialLogin();

for(var i=0;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
//Delay(20000);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
//Delay(5000);

todo(temp[3]);
aprvBudget(temp[0],temp[1],temp[2]);

//WorkspaceUtils.closeMaconomy();
//Delay(20000);
}
  
  
  
 // status = Login_Match();
  
//  if(status){
//  RestMaconomy()
///  gotoApprove()
//  }
  }
  WorkspaceUtils.closeAllWorkspaces();
}
}


function aprvBudget(ComId,EmpNo,userNmae){ 
aqUtils.Delay(5000, Indicator.Text);
// Delay(5000) 
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
//Delay(5000)
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);
//Delay(3000);
} 

// var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//   var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//   var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
//    SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
//    SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).
//    SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
//    SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
//    SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
   
    aqUtils.Delay(8000, Indicator.Text);
//    Delay(2000);

Delay(3000)
var All_User =
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
 //Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click(); 
  Delay(4000);
  var firstcell = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
 // firstcell.Keys("1307_AutomationUser 24September2019 16:20:30");
   firstcell.Keys(nameValue+" "+STIME);
  
//  firstcell.Keys("Muthu");
  firstcell.Keys("[Tab][Tab]");
  Delay(2000);
  var Employee_no = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2)
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
//  Employee_no.Keys("170710134");
  Employee_no.Keys(employeeNo);
  Delay(4000);
  
  var table = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==("1307_AutomationUser 24September2019 16:20:30"))
        if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue+" "+STIME)){  
      flag=true;
      break;
    }else{ 
      table.Keys("[Down]");
    }
  } 
  
     ValidationUtils.verify(flag,true,"Job is listed for Approval"); 
  if(flag){ 
   var closefilter = 
   Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
 //  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    closefilter.Click();
    Delay(5000);
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var approve =   Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8)
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8)
//  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
  Sys.HighlightObject(approve);

 if(approve.isEnabled()){ 
  approve.Click();
//ValidationUtils.verify(true,true,"Job is Approved by :"+userNmae)
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
//  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
}
  Delay(5000);
  var Ok = 
  Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
 // Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  //
  Ok.Click();
   Delay(2000);
  if(Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
    
  //Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
    var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  expDate.Click();
  }
  
  }
  
  
  
    
//    companyFilter.forceFocus();
//    companyFilter.setVisible(true);
//    companyFilter.ClickM();
//    table.Child(0).setText("^a[BS]");
//    table.Child(0).setText("1307_AutomationUser 24September2019 16:20:30");
//    aqUtils.Delay(2000, Indicator.Text);
////    Delay(2000);
//    table.Child(0).Keys("[Tab][Tab]");
//
//    var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//    
//    job.ClickM();
//    table.Child(2).forceFocus();
//    table.Child(2).setVisible(true);
//    table.Child(2).setText("^a[BS]");
//    table.Child(2).setText(EmpNo);
//    aqUtils.Delay(3000, Indicator.Text);
//    
//    
////    Delay(3000);
//var flag=false;
//for(var v=0;v<table.getItemCount();v++){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()==JobNo){ 
//  flag=true;
//  break;
//}
//else{ 
//  table.Keys("[Down]");
//}
//}
//    ValidationUtils.verify(flag,true,"Job is listed for Approval");
//    
//    if(table.getItemCount()>0){
////    Log.Message("Created Job is listed in table")
//    closeFilter.Click();
//    aqUtils.Delay(8000, Indicator.Text);
////    Delay(8000);
//    
////    var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
////    Budget.Click();
////    Delay(2000);
////    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
////    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
////    Sys.HighlightObject(show_budget);
////    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
//
////    var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
//    var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1307 Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
//    Budget.Click();
//    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
//    ref.Refresh();
//    aqUtils.Delay(5000, Indicator.Text);
////    Delay(5000);
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").Index==1)
//    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
//    
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==1)
//    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
//    Sys.HighlightObject(show_budget);
//    
//
//    show_budget.Keys("Working Estimate");
//    aqUtils.Delay(5000, Indicator.Text);
////    Delay(7000);
//    var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 9);
//Sys.HighlightObject(Approve)
//if(Approve.isEnabled()){ 
////  Approve.Click();
//ValidationUtils.verify(true,true,"Job is Approved by :"+userNmae)
//}
//else{ 
//  ReportUtils.logStep("INFO","Approve Button Is Invisible");
//  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
//}
//    }

}


function test()
{
 
if(Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
    
  
    var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  expDate.Click();
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
//  Delay(2000);
//  Sys.Desktop.KeyDown(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Delay(2000);
//  Sys.Desktop.KeyDown(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Delay(2000);
//  Sys.Desktop.KeyDown(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Delay(2000);
//  Sys.Desktop.KeyDown(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Delay(2000);
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
if(lvl==2)
Client_Managt.DblClickItem("|Approve User Information (*)");

if(lvl==3)
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








//function CreateUser(){ 
//STIME = WorkspaceUtils.StartTime();
//  goToMenu();
//  goToUsers();
//  
//  getDetails();
//  
// // userInfo =  excel(sheetName);
//  var status = UserInformation();
////var status = true;
//  if(status){
//  user_Approval();
////Approve_Level[0] = " Muthu 22:06:31*170710134*SACHINDRA P KARKERA (170710011)*";
//  status = Login_Match();
//  
//  if(status){
//  RestMaconomy()
////  gotoApprove()
//  }
//  }
//  WorkspaceUtils.closeAllWorkspaces();
//}