//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "AbsenceRequest";
var Language = "";
Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var firstday,lasyday,duration,absencetype,Reason1 = "";



//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);

duration = ExcelUtils.getRowDatas("Duration",EnvParams.Opco)
Log.Message(duration)
if((duration==null)||(duration=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Absence Request");
}

//ExcelUtils.setExcelName(workBook, "Data Management", true);
//absencetype = ReadExcelSheet("AllowanceAbsenceType",EnvParams.Opco,"Data Management");
//if((absencetype=="")||(absencetype==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
absencetype = ExcelUtils.getRowDatas("Absence Type",EnvParams.Opco)
//}
if((absencetype==null)||(absencetype=="")){ 
ValidationUtils.verify(false,true,"Absence Type is Needed to Create a Absence Request");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
firstday = ExcelUtils.getRowDatas("First Day",EnvParams.Opco)
//  if((firstday=="")||(firstday==null)){
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  firstday = ReadExcelSheet("AllowanceAbsenceDate",EnvParams.Opco,"Data Management");
//  }
if((firstday==null)||(firstday=="")){ 
ValidationUtils.verify(false,true,"First Day is Needed to Create a Absence Request");
}

Reason1 = ExcelUtils.getRowDatas("Reason",EnvParams.Opco)
Log.Message(Reason1)
if((Reason1=="")||(Reason1==null)){
ValidationUtils.verify(false,true,"Reason is Needed to Create a Absence Request");
}

}




function gotoAbsence() {
  ReportUtils.logStep("INFO", "Enter Payment File Details");
 var absence = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absence);
 waitForObj(absence);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
 var absencerequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(absencerequest);
 absencerequest.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

  }
  
  
var newrequest;
  if(Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.isVisible())
  newrequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl3;
 else
  newrequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);

                
// var newrequest = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl3;
 Sys.HighlightObject(newrequest)
 waitForObj(newrequest);
 ReportUtils.logStep_Screenshot();
  newrequest.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
aqUtils.Delay(5000, Indicator.Text);
  Log.Message(firstday);
  var newabsencerequest = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Absence Request").OleValue.toString().trim()); 
 Sys.HighlightObject(newabsencerequest);
 
var Firstday = Aliases.Maconomy.CreateAbsence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McDatePickerWidget;  
//Aliases.Maconomy.Absence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McDatePickerWidget;
Sys.HighlightObject(Firstday);
if(firstday!=""){
aqUtils.Delay(1000, Indicator.Text);
Firstday.setText(firstday)
ValidationUtils.verify(true,true,"First Date is selected in Maconomy"); 
}
else{ 
ValidationUtils.verify(false,true,"First Date is Needed to Create a Absence Request");
} 
     
var Duration = Aliases.Maconomy.CreateAbsence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McTextWidget;    
Sys.HighlightObject(Duration);
Duration.setText(duration)
    
//    var Lastday = Aliases.Maconomy.Absence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McDatePickerWidget;
//    Sys.HighlightObject(Lastday);
//    if(Lastday.getText()!=lasyday){
//      if(lasyday!=""){
//         aqUtils.Delay(1000, Indicator.Text);
//      WorkspaceUtils.CalenderDateSelection(Lastday,lasyday)
//      ValidationUtils.verify(true,true,"Last Date is selected in Maconomy"); 
//      }
//    }
//     else{ 
//        ValidationUtils.verify(false,true,"Last Date is Needed to Create a Absence Request");
//     } 
     
var type = Aliases.Maconomy.Absence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite4.McValuePickerWidget;
Sys.HighlightObject(type);
type.Click();
WorkspaceUtils.SearchByValue(type,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Type").OleValue.toString().trim(),absencetype,"Title");

      var FirstHalfDay = Aliases.Maconomy.Absence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite5.McPlainCheckboxView.Button;
      var LastHalfDay = Aliases.Maconomy.Absence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite6.McPlainCheckboxView.Buttonl;
           
      var remarks = Aliases.Maconomy.Absence.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite7.McTextWidget
      Sys.HighlightObject(remarks);
      remarks.setText(Reason1+" "+STIME);
      var Type = Reason1+" "+STIME;
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("AbsenceType",EnvParams.Opco,"Data Management",Type)

      TextUtils.writeLog("Absence Request details are Filled");
aqUtils.Delay(5000, Indicator.Text);
       var createbtn = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Absence Request").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());  
       Sys.HighlightObject(createbtn);
        if(createbtn.isEnabled()){   
          createbtn.HoverMouse();
          ReportUtils.logStep_Screenshot(""); 
          createbtn.Click();
          TextUtils.writeLog("Absence Request is CREATED");
          ValidationUtils.verify(true,true,"Absence Request is CREATED");
        } 
        else{
          var cancelbtn = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Absence Request").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());   
          Sys.HighlightObject(cancelbtn)    
          cancelbtn.HoverMouse();
          ReportUtils.logStep_Screenshot("");
          cancelbtn.Click();
          TextUtils.writeLog("Absence Request is not CREATED");
          ValidationUtils.verify(true,false,"Absence Request is not Created");
        } 
        
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
         var table = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 Sys.HighlightObject(table);
 waitForObj(table)
 
 var all = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
 Sys.HighlightObject(all);
 all.Click();
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
 Duration.setText(duration);
 Duration.Keys("[Tab][Tab][Tab][Tab][Tab]")
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
 var Remarks = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
 Sys.HighlightObject(Remarks);
 Remarks.Click();
 Remarks.setText(Type);
 aqUtils.Delay(6000, "Playback");  
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}
 aqUtils.Delay(6000, "Playback");  
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

}

  var flag=false;
  var Duration = "";
  for(var v=0;v<table.getItemCount();v++){ 
  Log.Message(table.getItem(v).getText_2(6).OleValue.toString().trim())
  Log.Message(Type)
  Log.Message(table.getItem(v).getText_2(6).OleValue.toString().trim()==Type)
  if(table.getItem(v).getText_2(6).OleValue.toString().trim()==Type){ 
    Duration = table.getItem(v).getText_2(1).OleValue.toString().trim()
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
 ValidationUtils.verify(flag,true,"Created Absence Request is available in system");
     
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
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Absence Reason",EnvParams.Opco,"Data Management",Type)  
ExcelUtils.WriteExcelSheet("Absence Type",EnvParams.Opco,"Data Management",absencetype)  
ExcelUtils.WriteExcelSheet("Absence Date",EnvParams.Opco,"Data Management",firstday)  
ExcelUtils.WriteExcelSheet("Absence Duration",EnvParams.Opco,"Data Management",Duration) 



TextUtils.writeLog("Absence Request is Submitted");
       
var submittedicon = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.Click();
        
var submitted = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();
Log.Message(submitted)
TextUtils.writeLog("Absence Request Submitted by:"+submitted);        
  
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
ReportUtils.logStep("INFO", "Moved to Time & Expenses from Absence Menu");
TextUtils.writeLog("Moved to Time & Expenses from Absence Menu");
}



//Main Function
function CreateAbsenceRequest() {
TextUtils.writeLog("Create Absence Request Started"); 
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
sheetName = "AbsenceRequest";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
firstday,lasyday,duration,absencetype,Reason1 = "";
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
//ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

//  try{
    getDetails();
    goToJobMenuItem();   
    gotoAbsence(); 
//  }
//    catch(err){
//      Log.Message(err);
//    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



