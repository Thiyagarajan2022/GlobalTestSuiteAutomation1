//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart


/** 
 * This script implements Creation of New User
 * @author  : Sai Kiran Vemula
 * @version : 1.0
 * Created Date :09/17/2020
 */
 
/** 
 * This script modified for pop-up issue
 * @author  : Sai Kiran Vemula
 * @version : 2.0
 * Created Date :03/21/2022
 */
 
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "UserCreation";
var name,company,ApproverGroup,userType,AbsenceApprover,validityPeriodFrom,validityPeriodTo,accessLevel,Language = "";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var employeeNo = "";
var UserName = "";

var login_satuts = true;
var LoginEmp = [];
var HRData = [];
var temp_user = [];


var UserLevel = [];
var User_Login = [];
var approvers;
var STIME="";
var Project_manager = "";


/**
  *  This Main function invokes maconomy and calls subfunctionality methods
  */
function CreateUser(){ 
TextUtils.writeLog("Create New User"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create User test started::"+STIME);
TextUtils.writeLog("Create User test started::"+STIME);
try{
goToMenu(); 
goToUsers();
getDetails();
user_Information();
approverDetails();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();

for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApproveUser(temp[2],i);
}

}
catch(err){ 
  Log.Message(err);
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}

/**
  *  This function Navigates to Users screen from Human Resourses workspace
  */

function goToMenu(){ 
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
  menuBar.DblClick();
Log.Message(Language) 
  if(Language == "Chinese (Simplified)"){ 
if(ImageRepository.ImageSet4.Emp_Chinese_1.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_1.Click();
}
else if(ImageRepository.ImageSet4.Emp_Chinese_2.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_2.Click();
}
else if(ImageRepository.ImageSet4.Emp_Chinese_3.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_3.Click();  
}
  }else if(Language == "Spanish"){ 
    
if(ImageRepository.ImageSet4.Emp_Spanish_1.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_1.Click();
}
else if(ImageRepository.ImageSet4.Emp_Spanish_2.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_2.Click();
}
else if(ImageRepository.ImageSet4.Emp_Spanish_3.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_3.Click();  
}

  }else{
if(ImageRepository.ImageSet.Human_Resource_1.Exists()){
ImageRepository.ImageSet.Human_Resource_1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();  
}
}

  aqUtils.Delay(3000, "Finding Users");
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, "Finding Users");
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

for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users").OleValue.toString().trim());

}

}
TextUtils.writeLog("Entering into Users from Human Resources Menu");
}


/**
  *  This function gets required data from Data sheet
  */
function goToUsers(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();

var new_user = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
waitForObj(new_user);
new_user.HoverMouse();
ReportUtils.logStep_Screenshot();
new_user.Click();

TextUtils.writeLog("New User is Clicked"); 
}

//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
name = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if((name==null)||(name=="")){ 
ValidationUtils.verify(false,true,"name is Needed to Create a User");
}

ExcelUtils.setExcelName(workBook,"Data Management", true);
employeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
if((employeeNo==null)||(employeeNo=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
employeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
}
Log.Message(employeeNo)
if((employeeNo==null)||(employeeNo=="")){ 
ValidationUtils.verify(false,true,"Employee Number is Needed to Create a User");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
company = ExcelUtils.getRowDatas("Company",EnvParams.Opco)
Log.Message(company)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"company Number is Needed to Create a User");
}

userType= ExcelUtils.getRowDatas("User Type",EnvParams.Opco)
if((userType==null)||(userType=="")){ 
ValidationUtils.verify(false,true,"User Type is Needed to Create a User");
}

validityPeriodFrom= ExcelUtils.getRowDatas("Valid From",EnvParams.Opco)
if((validityPeriodFrom==null)||(validityPeriodFrom=="")){ 
ValidationUtils.verify(false,true,"Valid Period From is Needed to Create a User");
}

validityPeriodTo= ExcelUtils.getRowDatas("Valid To",EnvParams.Opco)
if((validityPeriodTo==null)||(validityPeriodTo=="")){ 
ValidationUtils.verify(false,true,"Valid Period To is Needed to Create a User");
}

accessLevel= ExcelUtils.getRowDatas("Access Level",EnvParams.Opco)
if((accessLevel==null)||(accessLevel=="")){ 
ValidationUtils.verify(false,true,"Access Level is Needed to Create a User");
}
}

function user_Information(){ 
Log.Message("Entering User Details");
aqUtils.Delay(2000,"Waiting for Create User window");
var createUserWindow = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim())
waitForObj(createUserWindow);
Sys.HighlightObject(createUserWindow);
Log.Message("User Creation Window Opened");

UserName = name+" "+STIME
var nameObj = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
waitForObj(nameObj);
nameObj.Click();
nameObj.setText(UserName);
ValidationUtils.verify(true,true,"Employee Name is entered in Maconomy");
  
var employeeNoObj = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2)
waitForObj(employeeNoObj);
employeeNoObj.Click();
WorkspaceUtils.SearchByValue(employeeNoObj,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),employeeNo,"Employee Number");
  

var companyObj = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
companyObj.Click();
WorkspaceUtils.SearchByValue(companyObj,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),company,"Company Number");


var userTypeObj = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
userTypeObj.Click();
WorkspaceUtils.SearchByValue(userTypeObj,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "User Type").OleValue.toString().trim(),userType,"User Type");


var validityFrom = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 2);
if(validityPeriodFrom = "AUTOFILL")
{ validityPeriodFrom = getSpecificDate(0);
  validityFrom.setText(validityPeriodFrom)
}
else
WorkspaceUtils.CalenderDateSelection(validityFrom,validityPeriodFrom)
ValidationUtils.verify(true,true,"Validity Period From is selected in Maconomy"); 

var validityTo= Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 4);
if(validityPeriodTo = "AUTOFILL"){
validityPeriodTo = getSpecificDate(30);
validityTo.setText(validityPeriodTo)
}
else
WorkspaceUtils.CalenderDateSelection(validityTo,validityPeriodTo)
ValidationUtils.verify(true,true,"Validity Period To is selected in Maconomy"); 

var createButton = Aliases.Maconomy.SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create User").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Create").OleValue.toString().trim())
waitForObj(createButton);
Sys.HighlightObject(createButton);
createButton.HoverMouse();
ReportUtils.logStep_Screenshot();
createButton.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
TextUtils.writeLog("Details is entered in screen and clicked Create"); 
aqUtils.Delay(6000,"Waiting until all Popups loaded");


 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim())
 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim(), 2000) ;

if (w.length > 0)
for(var popup_Index=0;popup_Index<w.length;popup_Index++){ 
  if ((w[popup_Index].Exists) && (w[popup_Index].Enabled) )
{

var label = w[popup_Index].SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = w[popup_Index].SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
Ok.Click();
Delay(8000);

break;
}

}

 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim(), 2000) ;

if (w.length > 0)
for(var popup_Index=0;popup_Index<w.length;popup_Index++){ 
  if ((w[popup_Index].Exists) && (w[popup_Index].Enabled) )
{

var label = w[popup_Index].SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = w[popup_Index].SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
Ok.Click();
Delay(8000);

break;
}

}
 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim(), 2000) ;

if (w.length > 0)
for(var popup_Index=0;popup_Index<w.length;popup_Index++){ 
  if ((w[popup_Index].Exists) && (w[popup_Index].Enabled) )
{

var label = w[popup_Index].SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = w[popup_Index].SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
Ok.Click();
Delay(10000);

break;
}

}

 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim(), 2000) ;

if (w.length > 0)
for(var popup_Index=0;popup_Index<w.length;popup_Index++){ 
  if ((w[popup_Index].Exists) && (w[popup_Index].Enabled) )
{

var label = w[popup_Index].SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = w[popup_Index].SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
Ok.Click();
Delay(8000);

break;
}

}

 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim(), 2000) ;

if (w.length > 0)
for(var popup_Index=0;popup_Index<w.length;popup_Index++){ 
  if ((w[popup_Index].Exists) && (w[popup_Index].Enabled) )
{

var label = w[popup_Index].SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = w[popup_Index].SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
Ok.Click();
Delay(8000);

break;
}

}

 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - Users").OleValue.toString().trim(), 2000) ;

if (w.length > 0)
for(var popup_Index=0;popup_Index<w.length;popup_Index++){ 
  if ((w[popup_Index].Exists) && (w[popup_Index].Enabled) )
{

var label = w[popup_Index].SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = w[popup_Index].SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
Ok.Click();
Delay(8000);

break;
}

}


}

function approverDetails()
{ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }  

var blockedUsers = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blocked Users").OleValue.toString().trim())
waitForObj(blockedUsers)
Sys.HighlightObject(blockedUsers);
blockedUsers.Click();

var table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
waitForObj(table);

var firstCell = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
firstCell.setText(UserName);  
aqUtils.Delay(2000,"Waiting for results to filter");
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(UserName)){ 
    flag=true;
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
   
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"User Created is available in system");
TextUtils.writeLog("User Employee is available in Maconomy screen");
  
  
if(flag){ 
   var closefilter = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
   closefilter.Click();    
 
   
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ImageRepository.ImageSet.Maximize1.Click();

aqUtils.Delay(1000,"Waiting for octa fields to appear");

var roleInfoPane =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "")
waitForObj(roleInfoPane);
Sys.HighlightObject(roleInfoPane);

var octaUserName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2)
waitForObj(octaUserName);
Sys.HighlightObject(octaUserName);
octaUserName.setText(EnvParams.Country + Math.floor(Math.random() * 100));

var octaOrg =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2) 
waitForObj(octaOrg);
Sys.HighlightObject(octaOrg);
octaOrg.setText(EnvParams.Country +  Math.floor(Math.random() * 100));

var octaSave = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3)
Sys.HighlightObject(octaSave);
octaSave.Click()
 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var userinfoWindow = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "")
waitForObj(userinfoWindow);
userinfoWindow.Click()

aqUtils.Delay(1000,"Waiting for submit button to appear");

var submitBtn = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
waitForObj(submitBtn);
Sys.HighlightObject(submitBtn);
submitBtn.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var approve_Bar = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl
waitForObj(approve_Bar);
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Maximize.Click();

var All_approver = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
All_approver.Click();
ReportUtils.logStep_Screenshot();


var approver_table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
waitForObj(approver_table);

var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   approvers="";
   if(approver_table.getItem(z).getText_2(5)!="Approved"){
//   approvers = approver_table.getItem(z).getText_2(7).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(8).OleValue.toString().trim();
//   ReportUtils.logStep("INFO","Employee Approver level : " +z+ " Approver :" +approvers);
//   Approve_Level[y] = company+"*"+approvers;
   
       Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
           
      //Self Approve is Disabled. So finding different Approver
        var mainApprover = approver_table.getItem(z).getText_2(6).OleValue.toString().trim();
        var substitur = approver_table.getItem(z).getText_2(7).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+UserName+"*"+ temp;
      Log.Message("Approver level :" +z+ ": " +approvers);
      Approve_Level[y] = approvers;
      
   Log.Message(Approve_Level[y])
   y++;
   }
}
}
TextUtils.writeLog("Finding approvers for Created User");

var childcount = 0;
var Add = [];
var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
for(var i = 0;i<Parent.ChildCount;i++){ 
  if(Parent.Child(i).isVisible()){
  Add[childcount] = Parent.Child(i);
  childcount++;
  }
}

Parent = "";
var pos = 0;
for(var i=0;i<Add.length;i++){ 
  if(Add[i].Height>pos){ 
    pos = Add[i].Height;
    Parent = Add[i];
  }
}


Log.Message(Parent.FullName)
var info_Bar = Parent.SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");;
Sys.HighlightObject(info_Bar);
Log.Message(info_Bar.FullName)
info_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


}


function CredentialLogin(){ 
  var AppvLevl = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("UAE")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("UAE")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }

}





function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){
  }
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){
  }
  
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}
Sys.HighlightObject(Client_Managt);
var listPass = true;


if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve User Information from To-Dos List"); 
listPass = false; 
break;
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
//if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
  if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve User Information (Substitute) from To-Dos List");  
var listPass = false;   
break;
  }
} 


}


function FinalApproveUser(userNmae,apvLvl){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(ImageRepository.ImageSet.Close_Filter.Exists()){ }else{ 
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);  
}


var table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;

var firstCell = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Log.Message("Final approval function "+name)
firstCell.setText(name+" "+STIME);

var closefilter = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(name+" "+STIME)){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
} 
    
var msg ="not"
if (flag)
  msg =""  
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"User is" +msg+" available in system");
TextUtils.writeLog("User is" +msg+" available in Maconomy screen");
    
if(flag){ 
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var ApproveBtn = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(ApproveBtn)
if(ApproveBtn.isEnabled()){ 
ApproveBtn.HoverMouse();
ReportUtils.logStep_Screenshot();
  ApproveBtn.Click(); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  

ValidationUtils.verify(true,true,"Created Employee is Approved by :"+userNmae)
TextUtils.writeLog("Created Employee is Approved by :"+userNmae); 
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(company+"- Approver :"+userNmae);
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(apvLvl==(ApproveInfo.length-1)){

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
aqUtils.Delay(10000,"Maconomy loading data");
}

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
aqUtils.Delay(10000,"Maconomy loading data");
}



 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
aqUtils.Delay(10000,"Maconomy loading data");
}

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
aqUtils.Delay(10000,"Maconomy loading data");
}

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();

var approve_Bar = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
waitForObj(approve_Bar);
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var approvalsTab = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
approvalsTab.Click();
aqUtils.Delay(4000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ReportUtils.logStep_Screenshot();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var approver_table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
for(var i=0;i<approver_table.getItemCount();i++){   
if(approver_table.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(false,true,"User is NOT APPROVED in Level :"+i);
}
else{ 
ValidationUtils.verify(true,true,"User is APPROVED in Level :"+i);  
}
}
TextUtils.writeLog("User is Approved in all levels");   
  
var info_Bar = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel2.TabControl;
info_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Forward.Click();

var show_filter = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
show_filter.Click();


var activeUsersBtn = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Users").OleValue.toString().trim());
waitForObj(activeUsersBtn);
activeUsersBtn.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
waitForObj(table);
var firstCell = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
firstCell.setText(name+" "+STIME);  
aqUtils.Delay(2000,"Waiting until results filter");
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(name+" "+STIME)){ 
    flag=true;
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Created User is activated and available in system");
TextUtils.writeLog("Created User is activated and available in system"); 

ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("UserCreation_UserName",EnvParams.Opco,"Data Management",name +" "+STIME);
TextUtils.writeLog("UserCreation_UserName: "+name +" "+STIME);


}
}
}





