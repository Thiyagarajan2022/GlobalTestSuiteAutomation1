﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ChangeUser";
var userName,validityPeriodFrom,validityPeriodTo,Language = "";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var employeeNo = "";

var approvers;
var STIME="";
var Project_manager = "";



function changeUser(){ 
  
TextUtils.writeLog("Change Existing User Information"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco)
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
getDetails();
changeUser_Information();
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


//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook,"Data Management", true);
userName = ExcelUtils.getRowDatas("UserCreation_UserName",EnvParams.Opco)
if((userName==null)||(userName=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);  
userName = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
}
Log.Message(userName)
if((userName==null)||(userName=="")){ 
ValidationUtils.verify(false,true,"UserName is Needed to Change a User Information");
}

ExcelUtils.setExcelName(workBook, sheetName, true);  
validityPeriodFrom= ExcelUtils.getRowDatas("Valid From",EnvParams.Opco)
Log.Message(validityPeriodFrom)
if((validityPeriodFrom==null)||(validityPeriodFrom=="")){
ValidationUtils.verify(false,true,"Validity Period From is Needed to Change a User Information");
}

validityPeriodTo= ExcelUtils.getRowDatas("Valid To",EnvParams.Opco)
Log.Message(validityPeriodTo)
if((validityPeriodTo==null)||(validityPeriodTo=="")){ 
ValidationUtils.verify(false,true,"Validity Period To is Needed to Change a User Information");
}

}

function changeUser_Information()
{ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }  

var activeUsers = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Users").OleValue.toString().trim())
waitForObj(activeUsers)
Sys.HighlightObject(activeUsers);
activeUsers.Click();

var table = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
waitForObj(table);

var firstCell = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
firstCell.setText(userName);  
aqUtils.Delay(2000,"Waiting for results to filter");
var flag=false;
if(table.getItemCount()>=1)
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userName)){ 
    flag=true;
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
else
ValidationUtils.verify(flag,false,"User is not available in system");

if(!flag)
ValidationUtils.verify(flag,false,"User is not available in system");
   
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"User is available in system");
TextUtils.writeLog("User is available in Maconomy screen");
  
  
if(flag){ 
   var closefilter = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
   closefilter.Click();    
 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var periodFromObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;
waitForObj(periodFromObj);
CalenderDateSelection(periodFromObj,validityPeriodFrom);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var periodToObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget2;
waitForObj(periodToObj);
CalenderDateSelection(periodToObj,validityPeriodTo);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var save = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
waitForObj(save);
save.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

aqUtils.Delay(2000,"Waiting for Popup Window");

if((Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption== JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim() )){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
aqUtils.Delay(1000,"Waiting for Popup Window to disappear");
}

var submit = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 7);
waitForObj(submit);
submit.Click();;
aqUtils.Delay(6000,"Approval Info is loading")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var approve_Bar = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;

var Add_Visible8 = true;
while(Add_Visible8){
if(approve_Bar.isEnabled()){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
Add_Visible8 = false;
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
}
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Maximize.Click();

var All_approver = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
All_approver.Click();
ReportUtils.logStep_Screenshot();


var approver_table = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
waitForObj(approver_table);

var y=0;
for(var z=0;z<approver_table.getItemCount();z++){ 
   approvers="";
   if(approver_table.getItem(z).getText_2(6)!="Approved"){
   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
   ReportUtils.logStep("INFO","User Approver level : " +z+ " Approver :" +approvers);
   Approve_Level[y] = company+"*"+approvers;
   Log.Message(Approve_Level[y])
   y++;
   }
}
}
TextUtils.writeLog("Finding approvers for Changed User");

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
var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);

if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
level = 1;
var Approve = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 9);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(1000,"Waiting for  popup window")
}
ValidationUtils.verify(true,true,"Change User is Approved by :"+Project_manager)
TextUtils.writeLog("Change User is Approved by :"+Project_manager);
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
}


if(Approve_Level.length==1){
  
var activeUsers = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Users").OleValue.toString().trim())
waitForObj(activeUsers)
Sys.HighlightObject(activeUsers);
activeUsers.Click();

var table = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
waitForObj(table);

var firstCell = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
firstCell.setText(userName);  
aqUtils.Delay(2000,"Waiting for results to filter");
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userName)){ 
    flag=true;
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
   
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"User is available in system");
TextUtils.writeLog("User is available in Maconomy screen");

var periodFromObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2);
waitForObj(periodFromObj);
var changedPeriodFrom = periodFromObj.getText().OleValue.toString().trim();
if(changedPeriodFrom == validityPeriodFrom){
ValidationUtils.verify(flag,true,"Validity Period From Change successfully reflected in system");
TextUtils.writeLog("Validity Period From Change successfully reflected in system"); 
}
else
{
ValidationUtils.verify(flag,false,"Validity Period From Change is not reflected in system");
TextUtils.writeLog("Validity Period From Change is not reflected in system"); 
}        

var periodToObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 4);
waitForObj(periodToObj);
var changedPeriodTo = periodToObj.getText().OleValue.toString().trim();
if(changedPeriodTo == validityPeriodTo){
ValidationUtils.verify(flag,true,"Validity Period To Change successfully reflected in system");
TextUtils.writeLog("Validity Period To Change successfully reflected in system"); 
}
else
{
ValidationUtils.verify(flag,false,"Validity Period To Change is not reflected in system");
TextUtils.writeLog("Validity Period To Change is not reflected in system"); 
}     

}
}



function CredentialLogin(){ 

for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var Cred = Approve_Level[i].split("*");
  for(var j=1;j<3;j++){
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 
    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }
  if(temp.length!=0){
    temp = temp+"*"+j;
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
  break;
  }
  }
  if((temp=="")||(temp==null))
  Log.Error("User Name is Not available for level :"+i);
}
WorkspaceUtils.closeAllWorkspaces();
}




function todo(lvl){ 
    var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
  
  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);

if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }

Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");

if(lvl==3){
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim()+ "(*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim()+ "(*)");
}
if((lvl==1)||(lvl==2)){
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim()+" (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim()+" (*)");
}
break;
}
}
}
}


function FinalApproveUser(userNmae,apvLvl){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(ImageRepository.ImageSet.Close_Filter.Exists()){ }else{ 
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);  
}


var table = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;

var firstCell = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
firstCell.setText(userName);

var closefilter = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userName)){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
} 
    
    
ReportUtils.logStep_Screenshot();  

ValidationUtils.verify(flag,true,"Change User is available in system");
TextUtils.writeLog("Change User is available in Approver List");  
  
  
if(flag){ 
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var ApproveBtn = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl4;
waitForObj(ApproveBtn);
if(ApproveBtn.isEnabled()){ 
ApproveBtn.HoverMouse();
ReportUtils.logStep_Screenshot();
  ApproveBtn.Click(); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  

ValidationUtils.verify(true,true,"Change User is Approved by :"+userNmae)
TextUtils.writeLog("Change User is Approved by :"+userNmae); 
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(company+"- Approver :"+userNmae);
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(apvLvl==(ApproveInfo.length-1)){

aqUtils.Delay(3000,"Waiting for Popup Window");

if((Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption== JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim() )){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
aqUtils.Delay(1000,"Waiting for Next Popup Window");
}

if((Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption== JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim())){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
aqUtils.Delay(1000,"Waiting for Next Popup Window");
}


var approve_Bar = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
waitForObj(approve_Bar);
approve_Bar.HoverMouse();
ReportUtils.logStep_Screenshot();
approve_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var approvalsTab = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
approvalsTab.Click();
ReportUtils.logStep_Screenshot();

var approver_table = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Log.Message(approver_table.getItemCount());

for(var i=0;i<approver_table.getItemCount();i++){   

if(approver_table.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(false,true,"Level "+i+"Is not Approved");
}
}
TextUtils.writeLog("User is Approved in all levels");   
  
var info_Bar = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel2.TabControl;
info_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Forward.Click();

var show_filter = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
show_filter.Click();


var activeUsersBtn = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Users").OleValue.toString().trim());
waitForObj(activeUsersBtn);
activeUsersBtn.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
waitForObj(table);
var firstCell = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
firstCell.setText(userName);  
aqUtils.Delay(2000,"Waiting until results filter");
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userName)){ 
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

var periodFromObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;
waitForObj(periodFromObj);
var changedPeriodFrom = periodFromObj.getText().OleValue.toString().trim();
if(changedPeriodFrom == validityPeriodFrom){
ValidationUtils.verify(flag,true,"Validity Period From Change successfully reflected in system");
TextUtils.writeLog("Validity Period From Change successfully reflected in system"); 
}
else
{
ValidationUtils.verify(flag,false,"Validity Period From Change is not reflected in system");
TextUtils.writeLog("Validity Period From Change is not reflected in system"); 
}        

var periodToObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget2;
waitForObj(periodToObj);
var changedPeriodTo = periodToObj.getText().OleValue.toString().trim();
if(changedPeriodTo == validityPeriodTo){
ValidationUtils.verify(flag,true,"Validity Period To Change successfully reflected in system");
TextUtils.writeLog("Validity Period To Change successfully reflected in system"); 
}
else
{
ValidationUtils.verify(flag,false,"Validity Period To Change is not reflected in system");
TextUtils.writeLog("Validity Period To Change is not reflected in system"); 
}  
}
}
}

