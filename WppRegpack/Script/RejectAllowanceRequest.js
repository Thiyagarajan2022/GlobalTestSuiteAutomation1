//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
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
//var firstday,lasyday,duration,absencetype = "";
var AbsenceType,Entrydate,TimeRegistered,Reason1 = ""; 
var login =[];
var Approve_Level=[];
var level =0;

//getting data from datasheet

function getDetails(){

Indicator.PushText("Reading Data from Excel");
ExcelUtils.setExcelName(workBook, sheetName, true);
sheetName="ApproveAllowanceRequest";
  aqUtils.Delay(3000, Indicator.Text);

AbsenceType = ExcelUtils.getRowDatas("Absence Type",EnvParams.Opco)
Log.Message(AbsenceType)
  if((AbsenceType=="")||(AbsenceType==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  AbsenceType = ReadExcelSheet("Absence Type",EnvParams.Opco,"Data Management");
  Log.Message(AbsenceType)
  }
    
  Entrydate = ExcelUtils.getRowDatas("Entry Date",EnvParams.Opco)
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"Entry Date is Needed to Create a Client");
}
Log.Message(Entrydate)
  
TimeRegistered = ExcelUtils.getRowDatas("Valid Till",EnvParams.Opco)
Log.Message(TimeRegistered)
  if((TimeRegistered=="")||(TimeRegistered==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  TimeRegistered = ReadExcelSheet("Valid Till",EnvParams.Opco,"Data Management");
  Log.Message(TimeRegistered)
  }
//  
  Reason1 = ExcelUtils.getRowDatas("Reason",EnvParams.Opco)
Log.Message(Reason1)
  if((Reason1=="")||(Reason1==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  Reason = ReadExcelSheet("Reason",EnvParams.Opco,"Data Management");
  Log.Message(Reason)
  }
  
Indicator.PushText("Playback");

}
//function getDetails(){
//ExcelUtils.setExcelName(workBook, sheetName, true);
//firstday = ExcelUtils.getRowDatas("First Day",EnvParams.Opco)
//Log.Message(firstday)
//if((firstday==null)||(firstday=="")){ 
//ValidationUtils.verify(false,true,"First Day is Needed to Create a Absence Request");
//}
//duration = ExcelUtils.getRowDatas("Duration",EnvParams.Opco)
//Log.Message(duration)
//if((duration==null)||(duration=="")){ 
//ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Absence Request");
//}
//
//absencetype = ExcelUtils.getRowDatas("AbsenceType",EnvParams.Opco)
//if((absencetype==null)||(absencetype=="")){ 
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//absencetype = ReadExcelSheet("AbsenceType",EnvParams.Opco,"Data Management");
//}
//if((absencetype==null)||(absencetype=="")){ 
//ValidationUtils.verify(false,true,"Absence Type is Needed to Create a Absence Request");
//}
//}




function gotoAbsence() {
  ReportUtils.logStep("INFO", "Enter Payment File Details");
  
  var absenceallowance1 = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(absenceallowance1);
absenceallowance1.HoverMouse();
ReportUtils.logStep_Screenshot("");
absenceallowance1.Click();
aqUtils.Delay(3000, "waiting for new absence allowance");

var all = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
                                
WorkspaceUtils.waitForObj(all);
all.HoverMouse();
ReportUtils.logStep_Screenshot("");
all.Click();
TextUtils.writeLog("New Allowance Request is clicked");
aqUtils.Delay(3000, "Checking Labels");


var c_enterdate=Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(c_enterdate);
c_enterdate.HoverMouse();
ReportUtils.logStep_Screenshot("");
c_enterdate.Click();
c_enterdate.setText(Entrydate);
TextUtils.writeLog("Entering EnterDate :"+Entrydate);
aqUtils.Delay(3000, "Checking Labels");

var supervisorinfo = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.Click();
var Supervisor = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();
Log.Message(Supervisor)
    var  split = Supervisor.split(" ")
   login = split[0]+" "+split[1]+" "+split[2]
   Log.Message(login)
   aqUtils.Delay(1000,"waiting for window");


//WorkspaceUtils.waitForObj(supervisorinfo);
//supervisorinfo.HoverMouse();
//var mess= supervisorinfo.getText();
//Log.Message(mess);
//var split= mess.split(" ")
//login = split[0]+" "+split[1]+" "+split[2]
//Log.Message(login)
//ReportUtils.logStep_Screenshot("");
////Log.Message(supervisorinfo.getText());
//TextUtils.writeLog("Supervisor Information :"+login);
//aqUtils.Delay(3000, "Checking Labels");

//var  supervisoradres = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.Click();
//  
//  var Supervisor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.getText().OleValue.toString().trim();
////   Log.Message(Supervisor)
//    var  split = Supervisor.split(" ")
//   login = split[0]+" "+split[1]+" "+split[2]
//   Log.Message(login)
//   aqUtils.Delay(1000,"waiting for window");

}

  
// var absence = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(absence);
// waitForObj(absence);
// var absencerequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(absencerequest);
// absencerequest.Click();
// 
// var submitted = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
// Sys.HighlightObject(submitted)
//  submitted.Click();
//  
// var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
// Sys.HighlightObject(table);
// 
// 
// var Firstday = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
// Sys.HighlightObject(Firstday);
// Firstday.Click();
// Firstday.setText(firstday);
// Firstday.Keys("[Tab]")
// 
//  var Duration = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
//  Sys.HighlightObject(Duration);
// Duration.Click();
// Duration.setText(duration);
// Duration.Keys("[Tab][Tab][Tab][Tab][Tab]")
//    
// var Remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
// Sys.HighlightObject(Remarks);
// Remarks.Click();
// Remarks.setText(absencetype);
 
// var submitted = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText();
// var submitby = submitted.split("*");

// var  supervisoradres = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.Click();
//  
//  var Supervisor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.getText().OleValue.toString().trim();
////   Log.Message(Supervisor)
//    var  split = Supervisor.split(" ")
//   login = split[0]+" "+split[1]+" "+split[2]
//   Log.Message(login)
//   aqUtils.Delay(1000,"waiting for window");
//}

function Approve(){
  
var absenceapproval = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(absenceapproval)
  waitForObj(absenceapproval)
  absenceapproval.HoverMouse();
ReportUtils.logStep_Screenshot("");
absenceapproval.Click();
aqUtils.Delay(6000, "waiting for new absence allowance");

  var allowancerequest = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(allowancerequest)
  //allowancerequest.Click();
    allowancerequest.HoverMouse();
ReportUtils.logStep_Screenshot("");
allowancerequest.Click();
  aqUtils.Delay(6000, "waiting for new absence allowance");
  
  
  var Awaitingaprovaltab = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  Sys.HighlightObject(Awaitingaprovaltab);
  //var Awaitingaproval = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  Sys.HighlightObject(Awaitingaprovaltab);
  Awaitingaprovaltab.Click();
   Awaitingaprovaltab.HoverMouse();
ReportUtils.logStep_Screenshot("");
allowancerequest.Click();
  aqUtils.Delay(6000, "waiting for new absence allowance");
  
  
//  var table = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//  Sys.HighlightObject(table);
//  var employee = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
//  
////Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
//  //Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
//
//Sys.HighlightObject(employee);
//  employee.Click();
//  employee.Keys("[Tab][Tab]");
//  employee.setText(Entrydate);
//   aqUtils.Delay(6000, "waiting for new absence allowance");
//  var remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
//  Sys.HighlightObject(remarks);
//  remarks.Click();
//  remarks.setText(absencetype);
//  
//         Sys.Desktop.KeyDown(0x11);
//          Sys.Desktop.KeyDown(0x46);
//          Sys.Desktop.KeyUp(0x11);
//          Sys.Desktop.KeyUp(0x46); 
//          ReportUtils.logStep_Screenshot();
  
//  var approvetab = Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite
// //
//  Sys.HighlightObject(approvetab);
//   aqUtils.Delay(6000, "waiting for new absence allowance");
   
  var Approve =  Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  
//Aliases.Maconomy.AbsenceAllowance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(Approve)
  //allowancerequest.Click();
    //Approve.HoverMouse();
ReportUtils.logStep_Screenshot("");  
Approve.Click();
  TextUtils.writeLog("Absence Request is Approved by:"+login);  
  aqUtils.Delay(6000, "waiting for new absence allowance");
  
  
  var rejallowancerequest=Aliases.Maconomy.RejectAllowance;
  Sys.HighlightObject(rejallowancerequest);
  aqUtils.Delay(3000, "waiting for new absence allowance");
  
//  var valid =Aliases.Maconomy.ApproveAllowanceRequest.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McDatePickerWidget;
//  Sys.HighlightObject(valid);
//  ReportUtils.logStep_Screenshot("");  
//valid.Click();
//valid.setText(TimeRegistered);
//TextUtils.writeLog("Entering EnterDate :"+TimeRegistered);
//aqUtils.Delay(3000, "Checking Labels");
 

var rr1 = Aliases.Maconomy.RejectAllowance.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McTextWidget;  
Sys.HighlightObject(rr1);
  ReportUtils.logStep_Screenshot("");  
rr1.Click();
rr1.setText(Reason1);
TextUtils.writeLog("Entering EnterDate :"+Reason1);
aqUtils.Delay(3000, "Checking Labels");

var rejectrequest= Aliases.Maconomy.RejectAllowance.Composite.Composite.Composite2.Composite.Button;
 Sys.HighlightObject(rejectrequest);
  ReportUtils.logStep_Screenshot("");  
rejectrequest.Click();
 aqUtils.Delay(6000, "Checking Labels");
}

//Go To Job from Menu
//function goToJobMenuItem(){
//     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//      menuBar.HoverMouse();
//      ReportUtils.logStep_Screenshot("");
//    menuBar.DblClick();
//     if(ImageRepository.ImageSet0.TimeExpense.Exists()){
//       ImageRepository.ImageSet0.TimeExpense.Click();// GL
//      }
//     else if(ImageRepository.ImageSet0.TimeExpense1.Exists()){
//       ImageRepository.ImageSet0.TimeExpense1.Click();
//      }
//     else{
//       ImageRepository.ImageSet0.TimeExpense2.Click();
//    }
//
//var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
//var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//Delay(3000);
//var MainBrnch = "";
//for(var bi=0;bi<WrkspcCount;bi++){ 
//  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
//    MainBrnch = Workspc.Child(bi);
//    break;
//  }
//}
//
//
//var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
//  var Client_Managt;
////Log.Message(childCC)
//for(var i=1;i<=childCC;i++){ 
//Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
//if(Client_Managt.isVisible()){ 
//Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
//ReportUtils.logStep_Screenshot();
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
//}
//
//}    
//
//     aqUtils.Delay(5000, Indicator.Text);     
//     ReportUtils.logStep("INFO", "Moved to Time & Expenses from Absence Menu");
//     TextUtils.writeLog("INFO", "Moved to Time & Expenses from Absence Menu");
//}
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
Client_Managt.ClickItem("|Absence");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Absence");
}

} 

ReportUtils.logStep("INFO", "Moved to Absence from Absence Menu");
TextUtils.writeLog("Entering into Absence from Time & Expense Menu");
}



function CredentialLogin(){ 
        for(var i=level;i<login.length;i++){
            var UserN = true;
            var temp="";
            var Cred = login[i].split("*");
            Log.Message(Cred)
            for(var j=2;j<4;j++){
                if((Cred[j]!="")&&(Cred[j]!=null))
                    if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf("1307"+" ")!=-1)))
                    { 
                       var sheetName = "Agency Users";
                      ExcelUtils.setExcelName(workBook, sheetName, true);
                      temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
                    }
                    else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
                    { 
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
                     
                if(temp.length!=0){                
                  temp = temp+"*"+j;
                  ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;                  
//                  Log.Message(ApproveInfo[i]);       
                  logindetail[w] = temp;
                  w++;                                                  
                  break;
                }
            }
            if((temp=="")||(temp==null))
            Log.Error("User Name is Not available for level :"+i);
        }
        WorkspaceUtils.closeAllWorkspaces();
}


//Main Function
function ApproveAllowanceRequest() {
TextUtils.writeLog("Create Absence Request Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior APs","Username")
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
//    CredentialLogin();
    closeAllWorkspaces();
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

