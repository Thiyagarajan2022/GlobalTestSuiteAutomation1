//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ApproveAllowanceRequest";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var AbsenceType,Entrydate,TimeRegistered,Reason1 = ""; 
var login =[];
var Approve_Level=[];
var level =0;

//getting data from datasheet
function getDetails(){
Indicator.PushText("Reading Data from Excel");
sheetName="ApproveAllowanceRequest";
aqUtils.Delay(3000, Indicator.Text);

ExcelUtils.setExcelName(workBook, "Data Management", true);
AbsenceType = ReadExcelSheet("AllowanceAbsenceType",EnvParams.Opco,"Data Management");
if((AbsenceType=="")||(AbsenceType==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
AbsenceType = ExcelUtils.getRowDatas("Absence Type",EnvParams.Opco)
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Entrydate = ReadExcelSheet("AllowanceAbsenceDate",EnvParams.Opco,"Data Management");
if((Entrydate==null)||(Entrydate=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Entrydate = ExcelUtils.getRowDatas("Entry Date",EnvParams.Opco)
}
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Approve Allowance Request");
}
Log.Message(Entrydate)

ExcelUtils.setExcelName(workBook, sheetName, true);
TimeRegistered = ExcelUtils.getRowDatas("Valid Till",EnvParams.Opco)
if((TimeRegistered=="")||(TimeRegistered==null)){
ValidationUtils.verify(false,true,"Valid Till is Needed to Create a Approve Allowance Request");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Reason1 = ReadExcelSheet("Allowance Reason",EnvParams.Opco,"Data Management");
if((Reason1=="")||(Reason1==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Reason1 = ExcelUtils.getRowDatas("Reason",EnvParams.Opco)
}
if((Reason1=="")||(Reason1==null)){
ValidationUtils.verify(false,true,"Reason is Needed to Create a Approve Allowance Request");
}
  
Indicator.PushText("Playback");

}





function gotoAbsence() {
ReportUtils.logStep("INFO", "Enter Payment File Details");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var absenceallowance1 = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(absenceallowance1);
absenceallowance1.HoverMouse();
ReportUtils.logStep_Screenshot("");
absenceallowance1.Click();
aqUtils.Delay(3000, "waiting for new absence allowance");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//var all = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
var all = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim())
                                
WorkspaceUtils.waitForObj(all);
all.HoverMouse();
ReportUtils.logStep_Screenshot("");
all.Click();
TextUtils.writeLog("New Allowance Request is clicked");
aqUtils.Delay(3000, "Checking Labels");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var c_enterdate=Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(c_enterdate);
c_enterdate.HoverMouse();
ReportUtils.logStep_Screenshot("");
c_enterdate.Click();
c_enterdate.setText(Entrydate);
TextUtils.writeLog("Entering EnterDate :"+Entrydate);
aqUtils.Delay(3000, "Checking Labels");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

c_enterdate.Keys("[Tab][Tab][Tab]")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var Reson=Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget3
WorkspaceUtils.waitForObj(Reson);
Reson.HoverMouse();
ReportUtils.logStep_Screenshot("");
Reson.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Reson.setText(Reason1);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000, "Checking Labels");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000, "Checking Labels");
var table = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
waitForObj(table)
Sys.HighlightObject(table);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  Log.Message(table.getItem(v).getText_2(3).OleValue.toString().trim())
  Log.Message(Reason1)
  Log.Message(table.getItem(v).getText_2(3).OleValue.toString().trim()==Reason1)
  if(table.getItem(v).getText_2(3).OleValue.toString().trim()==Reason1){ 
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
 ValidationUtils.verify(flag,true,"Created Allowance Request is available in system");
 
 
var supervisorinfo = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).Click();
var Supervisor = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).getText().OleValue.toString().trim();
//var supervisorinfo = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.Click();
//var Supervisor = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();
Log.Message(Supervisor)
Log.Message(Supervisor.substring(0,Supervisor.lastIndexOf(" ")))
//    var  split = Supervisor.split(" ")
//   login = split[0]+" "+split[1]+" "+split[2]
login = Supervisor.substring(0,Supervisor.lastIndexOf(" "))
   Log.Message(login)
   aqUtils.Delay(1000,"waiting for window");

}

  


function Approve(){
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var absenceapproval = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(absenceapproval)
waitForObj(absenceapproval)
absenceapproval.HoverMouse();
ReportUtils.logStep_Screenshot("");
absenceapproval.Click();
aqUtils.Delay(6000, "waiting for new absence allowance");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var allowancerequest = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(allowancerequest)
allowancerequest.HoverMouse();
ReportUtils.logStep_Screenshot("");
allowancerequest.Click();
aqUtils.Delay(6000, "waiting for new absence allowance");
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
//  var Awaitingaprovaltab = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  var Awaitingaprovaltab = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Awaiting Approval").OleValue.toString().trim())
  Sys.HighlightObject(Awaitingaprovaltab);
  Sys.HighlightObject(Awaitingaprovaltab);
  Awaitingaprovaltab.Click();
  Awaitingaprovaltab.HoverMouse();
ReportUtils.logStep_Screenshot("");
allowancerequest.Click();
  aqUtils.Delay(6000, "waiting for new absence allowance");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 Management (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
  var table = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  var employee = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "");
  Sys.HighlightObject(employee);
  employee.Click();
  employee.Keys("[Tab][Tab]");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var entry = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "", 3);
entry.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
entry.setText(Entrydate);
aqUtils.Delay(6000, "waiting for new absence allowance");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  entry.Keys("[Tab][Tab][Tab][Tab]");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 Management (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
  var remarks = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "", 2);
  Sys.HighlightObject(remarks);
  remarks.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  remarks.setText(Reason1);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000, "Checking Labels");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  Log.Message(table.getItem(v).getText_2(6).OleValue.toString().trim())
  Log.Message(Reason1)
  Log.Message(table.getItem(v).getText_2(6).OleValue.toString().trim()==Reason1)
  if(table.getItem(v).getText_2(6).OleValue.toString().trim()==Reason1){ 
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
 ValidationUtils.verify(flag,true,"Created Allowance Request is available in system");
 ReportUtils.logStep_Screenshot();
 var closeFilter = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 closeFilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  

aqUtils.Delay(6000, "waiting for Reject");
  

   
var Approve =  Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(Approve)

Approve.HoverMouse();
ReportUtils.logStep_Screenshot("");  
Approve.Click();
TextUtils.writeLog("Absence Allowance Request is Approved by:"+login);  
aqUtils.Delay(6000, "waiting for new absence allowance");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
} 

  aqUtils.Delay(3000, "waiting for new absence allowance");
  
  var valid =Aliases.Maconomy.ApproveAllowanceRequest.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McDatePickerWidget;
  Sys.HighlightObject(valid);
  ReportUtils.logStep_Screenshot("");  
valid.Click();
valid.setText(TimeRegistered);
TextUtils.writeLog("Entering EnterDate :"+TimeRegistered);
aqUtils.Delay(3000, "Checking Labels");
 

var rr1 = Aliases.Maconomy.ApproveAllowanceRequest.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McTextWidget  
Sys.HighlightObject(rr1);
  ReportUtils.logStep_Screenshot("");  
rr1.Click();
rr1.setText("Approve");
TextUtils.writeLog("Entering EnterDate :"+Reason1);
aqUtils.Delay(3000, "Checking Labels");

//Spanish
var approverequest=Aliases.Maconomy.ApproveAllowanceRequest.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Request").OleValue.toString().trim());
 Sys.HighlightObject(approverequest);
  ReportUtils.logStep_Screenshot("");  
approverequest.Click();
aqUtils.Delay(6000, "Checking Labels");
ValidationUtils.verify(true, true,"Absence Allowance Request is Approved by:"+login);  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var Status = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
Status.Click();
Delay(2000)
Status = Status.getText().OleValue.toString().trim();

if(Status.indexOf(login)!=-1)
ValidationUtils.verify(true, true,"Absence Allowance Request is Rejected by:"+login); 
else
ValidationUtils.verify(true, false,"Absence Allowance Request is NOT Rejected by:"+login);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

}

//Go To Job from Menu
function goToJobMenuItem(){


var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.TimeExpense.Exists()){
 ImageRepository.ImageSet.TimeExpense.Click();// GL
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
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
}

} 

ReportUtils.logStep("INFO", "Moved to Absence from Absence Menu");
TextUtils.writeLog("Entering into Absence from Time & Expense Menu");
}



//function CredentialLogin(){ 
//        for(var i=level;i<login.length;i++){
//            var UserN = true;
//            var temp="";
//            var Cred = login[i].split("*");
//            Log.Message(Cred)
//            for(var j=2;j<4;j++){
//                if((Cred[j]!="")&&(Cred[j]!=null))
//                    if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf("1307"+" ")!=-1)))
//                    { 
//                       var sheetName = "Agency Users";
//                      ExcelUtils.setExcelName(workBook, sheetName, true);
//                      temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//                    }
//                    else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
//                    { 
//                      var sheetName = "SSC Users";
//                      ExcelUtils.setExcelName(workBook, sheetName, true);
//                      temp = ExcelUtils.SSCLogin(Cred[j],"Username");  
//                    }
//                    else{ 
//                     var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
//                      if(UserN){ 
//                        goToHR();
//                        UserN = false;
//                      }
//                      temp = searchNumber(Eno);
//                    }
//                     
//                if(temp.length!=0){                
//                  temp = temp+"*"+j;
//                  ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;                  
////                  Log.Message(ApproveInfo[i]);       
//                  logindetail[w] = temp;
//                  w++;                                                  
//                  break;
//                }
//            }
//            if((temp=="")||(temp==null))
//            Log.Error("User Name is Not available for level :"+i);
//        }
//        WorkspaceUtils.closeAllWorkspaces();
//}


//Main Function
function ApproveAllowanceRequest() {
TextUtils.writeLog("Approve Allowance Request Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ApproveAllowanceRequest";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

try{
getDetails();
goToJobMenuItem(); 
gotoAbsence(); 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
WorkspaceUtils.closeMaconomy();
Restart.login(login);
goToAbsenceMenuItem();
Approve();
closeAllWorkspaces();
}
catch(err){
Log.Message(err);
}
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


function goToAbsenceMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
    menuBar.DblClick();
     if(ImageRepository.ImageSet0.TimeExpense.Exists()){
       ImageRepository.ImageSet0.TimeExpense.Click();// GL
      }
     else if(ImageRepository.ImageSet0.TimeExpense1.Exists()){
       ImageRepository.ImageSet0.TimeExpense1.Click();
      }
     else{
       ImageRepository.ImageSet0.TimeExpense2.Click();
    }

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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Approval").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Approval").OleValue.toString().trim());
}

}    

     aqUtils.Delay(5000, Indicator.Text);     
     TextUtils.writeLog("INFO", "Moved to Time & Expenses from Absence Approval Menu");
}

