//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ApproveAbsence";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var firstday,lasyday,duration,absencetype = "";
var login =[];
var Approve_Level=[];
var level =0;

//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
firstday = ExcelUtils.getRowDatas("First Day",EnvParams.Opco)
Log.Message(firstday)
if((firstday==null)||(firstday=="")){ 
ValidationUtils.verify(false,true,"First Day is Needed to Create a Absence Request");
}
duration = ExcelUtils.getRowDatas("Duration",EnvParams.Opco)
Log.Message(duration)
if((duration==null)||(duration=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Absence Request");
}

absencetype = ExcelUtils.getRowDatas("AbsenceType",EnvParams.Opco)
if((absencetype==null)||(absencetype=="")){ 
  ExcelUtils.setExcelName(workBook, "Data Management", true);
absencetype = ReadExcelSheet("AbsenceType",EnvParams.Opco,"Data Management");
}
if((absencetype==null)||(absencetype=="")){ 
ValidationUtils.verify(false,true,"Absence Type is Needed to Create a Absence Request");
}
}




function gotoAbsence() {
  ReportUtils.logStep("INFO", "Enter Payment File Details");
 var absence = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absence);
 waitForObj(absence);
 var absencerequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absencerequest);
 absencerequest.Click();
 
 var submitted =  Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submitted").OleValue.toString().trim())
// Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
 Sys.HighlightObject(submitted)
  submitted.Click();
  
 var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 Sys.HighlightObject(table);
 waitForObj(table)
 
 var Firstday = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
 Sys.HighlightObject(Firstday);
 Firstday.Click();
 Firstday.setText(firstday);
 Firstday.Keys("[Tab]")
 
  var Duration = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  Sys.HighlightObject(Duration);
 Duration.Click();
 Duration.setText(duration);
 Duration.Keys("[Tab][Tab][Tab][Tab][Tab]")
    
 var Remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
 Sys.HighlightObject(Remarks);
 Remarks.Click();
 Remarks.setText(absencetype);
 
  var flag =false; 
        for(var i=0;i<table.getItemCount();i++){          
          if(table.getItem(i).getText_2(6).OleValue.toString().trim()==absencetype){
            flag = true;        
            break;
          }  
          else{
              table.Keys("[Down]");
          } 
        } 
        aqUtils.Delay(3000,Indicator.Text); 
        ReportUtils.logStep_Screenshot();    
        ValidationUtils.verify(flag,true,"Absence Type is available in system");
        TextUtils.writeLog("Absence Type"+ absencetype+" is available in system");
// var submitted = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText();
// var submitby = submitted.split("*");

 var  supervisoradres = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.Click();
  
  var Supervisor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.getText().OleValue.toString().trim();
//   Log.Message(Supervisor)
    var  split = Supervisor.split(" ")
   login = split[0]+" "+split[1]+" "+split[2]
   Log.Message(login)
    Approve_Level = login;
    Log.Message(Approve_Level);
   aqUtils.Delay(1000,"waiting for window");
}

function Approve(){
  var absencapp = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(absencapp)
  waitForObj(absencapp)
  var absencereq = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(absencereq)
  absencereq.Click();
  var Awaitingaprovaltab = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer;
  Sys.HighlightObject(Awaitingaprovaltab);
  var Awaitingaproval = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Awaiting Approval").OleValue.toString().trim());
  Sys.HighlightObject(Awaitingaproval);
  Awaitingaproval.Click();
  
  var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  var employee = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  Sys.HighlightObject(employee);
  employee.Click();
  employee.Keys("[Tab][Tab][Tab][Tab][Tab]");
  
  var remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;  
//Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(remarks);
  remarks.Click();
  remarks.setText(absencetype);
        var  column = table.getColumnCount();
        var row = table.getItemCount()
        
        for(var i=0;i<row;i++){
          if(table.getItem(i).getText(5).OleValue.toString().trim()==absencetype){
            ValidationUtils.verify(true,true,"Absence Type is available in the table");
            break;
          }
          else{
            table.Keys("[Down]");
          }
        } 
  
         Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46); 
          ReportUtils.logStep_Screenshot();
  
  var approvetab = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
  Sys.HighlightObject(approvetab);
  var Approve = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  ReportUtils.logStep_Screenshot("");
  waitForObj(Approve)
  Approve.Click();
  TextUtils.writeLog("Approve Absence Request is Approved by:"+login);  
  
}

//Go To Job from Menu
function goToJobMenuItem(){
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence").OleValue.toString().trim());
}

}    

     aqUtils.Delay(5000, Indicator.Text);     
//     ReportUtils.logStep("INFO", "Moved to Time & Expenses from Absence Menu");
     TextUtils.writeLog("Moved to Time & Expenses from Absence Menu");
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
     TextUtils.writeLog("Moved to Time & Expenses from Absence Approval Menu");
}


function CredentialLogin(){ 
  Log.Message(Approve_Level)
        for(var i=level;i<Approve_Level.length;i++){
            var UserN = true;
            var temp="";
            var Cred = Approve_Level[i].split("*");
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
function ApproveAbsence() {
TextUtils.writeLog("Approve Absence Request Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ApproveAbsence";

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
  }
    catch(err){
      Log.Message(err);
    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

