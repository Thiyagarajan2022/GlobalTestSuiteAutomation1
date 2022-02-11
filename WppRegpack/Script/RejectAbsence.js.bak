//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "RejectAbsence";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var firstday,lasyday,duration,absencetype,Reason1 = "";

var login =[];
var Approve_Level=[];
var level =0;

//getting data from datasheet
function getDetails(){

ExcelUtils.setExcelName(workBook, "Data Management", true);
firstday = ReadExcelSheet("Absence Date",EnvParams.Opco,"Data Management");
if((firstday==null)||(firstday=="")){  
ExcelUtils.setExcelName(workBook, sheetName, true);
firstday = ExcelUtils.getRowDatas("First Day",EnvParams.Opco)
}
if((firstday==null)||(firstday=="")){ 
ValidationUtils.verify(false,true,"First Day is Needed to Reject a Absence Request");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
duration = ReadExcelSheet("Absence Duration",EnvParams.Opco,"Data Management");
if((duration==null)||(duration=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
duration = ExcelUtils.getRowDatas("Duration",EnvParams.Opco)
}
if((duration==null)||(duration=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Reject a Absence Request");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
absencetype = ReadExcelSheet("Absence Type",EnvParams.Opco,"Data Management");
if((absencetype==null)||(absencetype=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
absencetype = ExcelUtils.getRowDatas("AbsenceType",EnvParams.Opco)
}
if((absencetype==null)||(absencetype=="")){ 
ValidationUtils.verify(false,true,"Absence Type is Needed to Reject a Absence Request");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Reason1 = ReadExcelSheet("Absence Reason",EnvParams.Opco,"Data Management");
if((Reason1==null)||(Reason1=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Reason1 = ExcelUtils.getRowDatas("AbsenceType",EnvParams.Opco)
}
if((Reason1==null)||(Reason1=="")){ 
ValidationUtils.verify(false,true,"Absence Type is Needed to Reject a Absence Request");
}


}




function gotoAbsence() {
  ReportUtils.logStep("INFO", "Enter Reject Absence Details");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var absence = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absence);
 waitForObj(absence);
 var absencerequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absencerequest);
 absencerequest.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var submitted =  Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submitted").OleValue.toString().trim())
 Sys.HighlightObject(submitted)
  submitted.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 Sys.HighlightObject(table);
 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var Firstday = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
 Sys.HighlightObject(Firstday);
 Firstday.Click();
 Firstday.setText(firstday);
 Firstday.Keys("[Tab]")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var Duration = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  Sys.HighlightObject(Duration);
 Duration.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 Duration.setText(duration);
 Duration.Keys("[Tab][Tab][Tab][Tab][Tab]")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var Remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
 Sys.HighlightObject(Remarks);
 Remarks.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 Remarks.setText(Reason1);
 ReportUtils.logStep_Screenshot("");
aqUtils.Delay(1000,"waiting for window");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(1000,"waiting for window");
    var  column = table.getColumnCount();
        var row = table.getItemCount()
        for(var i=0;i<row;i++){
          if(table.getItem(i).getText(6).OleValue.toString().trim()==Reason1){
            ValidationUtils.verify(true,true,"Absence Type is available in the table");
            TextUtils.writeLog("Absence Type is available in the table:"+Reason1);
            break;
          }
          else{
            table.Keys("[Down]");
          }
        } 
 
        
 var  supervisoradres = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).Click();

  var Supervisor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2).getText().OleValue.toString().trim();
// var  supervisoradres = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.Click();
//
//  var Supervisor = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Composite.McTextWidget.getText().OleValue.toString().trim();

login = Supervisor.substring(0,Supervisor.lastIndexOf(" "))
   Log.Message(login)
   aqUtils.Delay(1000,"waiting for window");
}

function Reject(){
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
  employee.Keys("[Tab][Tab]");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var day = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(day);
  day.Click();
  day.setText(firstday);
  day.Keys("[Tab]");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var A_duration = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  Sys.HighlightObject(A_duration);
  A_duration.Click();
  A_duration.setText(duration);
  A_duration.Keys("[Tab][Tab][Tab]");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var A_type = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  Sys.HighlightObject(A_type);
  A_type.Click();
  A_type.setText(absencetype);
  A_type.Keys("[Tab][Tab][Tab]");
aqUtils.Delay(1000,"waiting for window");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(1000,"waiting for window");
  var reason = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2
  Sys.HighlightObject(reason);
  reason.Click();
  reason.setText(Reason1);
aqUtils.Delay(1000,"waiting for window");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
        var  column = table.getColumnCount();
        var row = table.getItemCount()
        for(var i=0;i<row;i++){
          if(table.getItem(i).getText(9).OleValue.toString().trim()==Reason1){
            ValidationUtils.verify(true,true,"Absence Type is available in the table");
            break;
          }
          else{
            table.Keys("[Down]");
          }
        } 
  
//         Sys.Desktop.KeyDown(0x11);
//          Sys.Desktop.KeyDown(0x46);
//          Sys.Desktop.KeyUp(0x11);
//          Sys.Desktop.KeyUp(0x46); 
          ReportUtils.logStep_Screenshot();
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
closefilter.Click();
aqUtils.Delay(1000,"waiting for window");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2
  var rejecttab = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
  Sys.HighlightObject(rejecttab);
  var reject = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
  ReportUtils.logStep_Screenshot("");
  reject.Click();
aqUtils.Delay(1000,"waiting for window");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var rejectreq = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject Absence Request").OleValue.toString().trim());
  Sys.HighlightObject(rejectreq);
  var rejectreason = Aliases.Maconomy.AbsenceReject.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McTextWidget;
  Sys.HighlightObject(rejectreason);
  rejectreason.Click();
  rejectreason.setText("Reject");
  aqUtils.Delay(1000,"Waiting for window");
  var reject = Aliases.Maconomy.Absence.Composite.Composite.Composite2.Composite.Button2;
  Sys.HighlightObject(reject);
  reject.Click();
  TextUtils.writeLog("Reject Absence Request is Rejected by:"+login); 
  aqUtils.Delay(1000,"Waiting for window");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var Status_Check = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite3.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  Status_Check = Status_Check.getText();

if(Status_Check=="Reject")
ValidationUtils.verify(true, true,"Absence Request is Rejected by:"+login); 
else
ValidationUtils.verify(true, false,"Absence Request is NOT Rejected by:"+login);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}


function UnDoReject() {
  ReportUtils.logStep("INFO", "Enter Reject Absence Details");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var absence = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absence);
 waitForObj(absence);
 var absencerequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absencerequest);
 absencerequest.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var submitted =  Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim())
 Sys.HighlightObject(submitted)
  submitted.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 Sys.HighlightObject(table);
 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var Firstday = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
 Sys.HighlightObject(Firstday);
 Firstday.Click();
 Firstday.setText(firstday);
 Firstday.Keys("[Tab]")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var Duration = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  Sys.HighlightObject(Duration);
 Duration.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 Duration.setText(duration);
 Duration.Keys("[Tab][Tab][Tab][Tab][Tab]")
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var Remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
 Sys.HighlightObject(Remarks);
 Remarks.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 Remarks.setText(Reason1);
 ReportUtils.logStep_Screenshot("");
aqUtils.Delay(1000,"waiting for window");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(1000,"waiting for window");
    var  column = table.getColumnCount();
        var row = table.getItemCount()
        for(var i=0;i<row;i++){
          if(table.getItem(i).getText(6).OleValue.toString().trim()==Reason1){
            ValidationUtils.verify(true,true,"Absence Type is available in the table");
            TextUtils.writeLog("Absence Type is available in the table:"+Reason1);
            break;
          }
          else{
            table.Keys("[Down]");
          }
        } 
 
 var Undo;
  if(Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.isVisible())
  Undo = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SWTObject("SingleToolItemControl", "", 7);
 else
  Undo = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);

//       var Undo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 7);
       Sys.HighlightObject(Undo);
       ReportUtils.logStep_Screenshot();
       waitForObj(Undo);
       Undo.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
   aqUtils.Delay(1000,"waiting for window");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}

 var submit;
  if(Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.isVisible())
  submit = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl4;
 else
  submit = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);

//       var submit = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl4;
       Sys.HighlightObject(submit);
       ReportUtils.logStep_Screenshot();
       waitForObj(submit);
       submit.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
   aqUtils.Delay(1000,"waiting for window");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
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
function RejectAbsence() {
TextUtils.writeLog("Reject Absence Request Started"); 
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
sheetName = "RejectAbsence";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
firstday,lasyday,duration,absencetype,Reason1 = "";
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
Reject();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
goToJobMenuItem(); 
UnDoReject(); 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

  }
    catch(err){
      Log.Message(err);
    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



