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
FinalApproveUser(temp[1],temp[2],i);
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
if(validityPeriodFrom == "AUTOFILL")
  validityPeriodFrom = getSpecificDate(1)
Log.Message(validityPeriodFrom)
if((validityPeriodFrom==null)||(validityPeriodFrom=="")){
ValidationUtils.verify(false,true,"Validity Period From is Needed to Change a User Information");
}

validityPeriodTo= ExcelUtils.getRowDatas("Valid To",EnvParams.Opco)
if(validityPeriodTo == "AUTOFILL")
  validityPeriodTo = getSpecificDate(10)
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
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(1000,"Loading")
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
var msg ="not"
if (flag)
  msg =""  
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"User is" +msg+" available in system");
TextUtils.writeLog("User is" +msg+" available in Maconomy screen");
  
  
if(flag){ 
   var closefilter = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
   closefilter.Click();    
 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var periodFromObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget;
waitForObj(periodFromObj);
periodFromObj.setText(validityPeriodFrom)
//CalenderDateSelection(periodFromObj,validityPeriodFrom);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var periodToObj = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McDatePickerWidget2;
waitForObj(periodToObj);
periodToObj.setText(validityPeriodTo)
//CalenderDateSelection(periodToObj,validityPeriodTo);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var save = Aliases.Maconomy.ChangeUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
waitForObj(save);
save.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

aqUtils.Delay(5000,"Waiting for Popup Window");

//if((Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption== JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim() )){
//var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim(),1).SWTObject("Label", "*").getText();
//Log.Message(label1);
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim(),1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
//Ok.HoverMouse();
//ReportUtils.logStep_Screenshot();
//Ok.Click();
//aqUtils.Delay(1000,"Waiting for Popup Window to disappear");
//}


 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim(), 2000);
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
//   approvers = approver_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
//   ReportUtils.logStep("INFO","User Approver level : " +z+ " Approver :" +approvers);
//   Approve_Level[y] = company+"*"+approvers;

       Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
           
      //Self Approve is Disabled. So finding different Approver
        var mainApprover = approver_table.getItem(z).getText_2(3).OleValue.toString().trim();
        var substitur = approver_table.getItem(z).getText_2(4).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+userName+"*"+ temp;
      Log.Message("Approver level :" +z+ ": " +approvers);
      Approve_Level[y] = approvers;
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



/*

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

*/




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


function FinalApproveUser(userNmae,Approver,apvLvl){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(ImageRepository.ImageSet.Close_Filter.Exists()){ }else{ 
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);  
}


var table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;

var firstCell = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
firstCell.setText(userNmae);

var closefilter = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userNmae)){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
} 
    
    
ReportUtils.logStep_Screenshot();  

ValidationUtils.verify(flag,true,"User Changed is available in system");
TextUtils.writeLog("Changed User is available in Approver List");  
  
  
if(flag){ 
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var ApproveBtn = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(ApproveBtn)
if(ApproveBtn.isEnabled()){ 
ApproveBtn.HoverMouse();
ReportUtils.logStep_Screenshot();
  ApproveBtn.Click(); 
  aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  

ValidationUtils.verify(true,true,"Changed User is Approved by :"+Approver)
TextUtils.writeLog("Changed User is Approved by :"+Approver); 
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(company+"- Approver :"+Approver);
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(apvLvl==(ApproveInfo.length-1)){


aqUtils.Delay(5000,"Maconomy loading data");



for(var h=0;h<2;h++){
  var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information (Substitute)").OleValue.toString().trim(), 2000) ;

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
aqUtils.Delay(10000,"Maconomy loading data");

break;
}

}
}



for(var h=0;h<2;h++){
  var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
 var w = p.FindAllChildren("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve User Information").OleValue.toString().trim(), 2000) ;

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
aqUtils.Delay(10000,"Maconomy loading data");

break;
}

}
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
  Log.Message(approver_table.getItemCount());

for(var i=0;i<approver_table.getItemCount();i++){   

if(approver_table.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(false,true,"Changed User NOT APPROVED in Level :"+i);
}
else{ 
ValidationUtils.verify(true,true,"Changed User APPROVED in Level :"+i);  
}
}
TextUtils.writeLog("User is Approved in all levels");   
  
var info_Bar = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel2.TabControl;
info_Bar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Forward.Click();


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

var show_filter = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
show_filter.Click();


var activeUsersBtn = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Users").OleValue.toString().trim());
waitForObj(activeUsersBtn);
activeUsersBtn.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
waitForObj(table);
var firstCell = Aliases.Maconomy.NewUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
firstCell.setText(userNmae);  
aqUtils.Delay(2000,"Waiting until results filter");
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userNmae)){ 
    flag=true;
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Changed User is activated and available in system");
TextUtils.writeLog("Changed User is activated and available in system"); 


ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Changed_UserName",EnvParams.Opco,"Data Management",userNmae);
TextUtils.writeLog("Changed_UserName: "+userNmae);


}
}
}





