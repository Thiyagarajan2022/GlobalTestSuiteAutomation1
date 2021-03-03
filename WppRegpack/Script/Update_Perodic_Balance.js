//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


/**
 * This script to Update Perodic Balance for Absence
 * @author  : Muthu Kumar M
 * @version : 1.0
 * Created Date :03/01/2021
*/

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Update Perodic Balance";

//Global Variable
var Project_manager="";
var level =0;
var STIME = "";
var Language = "";
var Vacation_Calander,Period,Absence_Type,Employee_No,Total_Allowance = "";



//Main Function
function Updating_Perodic_Balance() {
  
TextUtils.writeLog("Updating Perodic Balance Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Updating Perodic Balance script
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
sheetName = "Update Perodic Balance";


ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
level =0;



Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);



getDetails();

aqUtils.Delay(5000, Indicator.Text);
goTo_Timeand_Expense_Item();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
Select_Vacation_Calander()
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
Change_Allowance_Period();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
Submit_and_Approve_Allowance();


var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}



//Getting Data from Datasheets
function getDetails(){ 
  
Vacation_Calander,Period,Absence_Type,Employee_No,Total_Allowance = "";
  ExcelUtils.setExcelName(workBook, "Update Perodic Balance", true);
  Vacation_Calander = ExcelUtils.getRowDatas("Vacation Calander No",EnvParams.Opco);
  Period = ExcelUtils.getRowDatas("Period",EnvParams.Opco);
  Absence_Type = ExcelUtils.getRowDatas("Absence Type",EnvParams.Opco);
  Employee_No = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco);
  Total_Allowance = ExcelUtils.getRowDatas("Total Allowance",EnvParams.Opco);
  
  if((Vacation_Calander=="")||(Vacation_Calander==null)){
ValidationUtils.verify(true,false,"Vacation Calander No is needed to update periodic Balance")
  }  

    if((Period=="")||(Period==null)){
ValidationUtils.verify(true,false,"Period is needed to update periodic Balance")
  }  
  
    if((Absence_Type=="")||(Absence_Type==null)){
ValidationUtils.verify(true,false,"Absence Type is needed to update periodic Balance")
  }  
  
    if((Employee_No=="")||(Employee_No==null)){
ValidationUtils.verify(true,false,"Employee No is needed to update periodic Balance")
  }  
  
      if((Total_Allowance=="")||(Total_Allowance==null)){
ValidationUtils.verify(true,false,"Total Allowance Period is needed to update periodic Balance")
  }  

}



/**
  *  This function Navigates to Absence Administration screen from Time and Expenses workspace
  */
function goTo_Timeand_Expense_Item(){
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Administration").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Administration").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Absence Administration from Time and Expenses Menu");
TextUtils.writeLog("Entering into Absence Administration from Time and Expenses Menu");
}



//Selecting Job for  Job Crediting in Maconomy
function Select_Vacation_Calander(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var Periodic_Balance = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(Periodic_Balance);
Periodic_Balance.Click();

aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var Calendar_No = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Calendar_No.Click();
Calendar_No.Keys(Vacation_Calander);

aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var CloseFilter = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
var table = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;

  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==Vacation_Calander){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
    CloseFilter.Click();
    aqUtils.Delay(3000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    
    var Current_Period = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
    Sys.HighlightObject(Current_Period);
    Current_Period.Click();
    WorkspaceUtils.VacationPeriod(Current_Period,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vacation Period").OleValue.toString().trim(),Period,Vacation_Calander,"Current Period");
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var AbsenceType = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
    Sys.HighlightObject(AbsenceType);
    AbsenceType.Click();
    WorkspaceUtils.SearchByValue(AbsenceType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Type").OleValue.toString().trim(),Absence_Type,"Absence Type");
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var Emp_No = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McTextWidget;
    Sys.HighlightObject(Emp_No);
    Emp_No.Click();
    Emp_No.Keys(Employee_No);
    
    var Emp_No = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McTextWidget2;
    Sys.HighlightObject(Emp_No);
    Emp_No.Click();
    Emp_No.Keys(Employee_No);
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var Company_No = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McTextWidget;
    Sys.HighlightObject(Company_No);
    Company_No.Click();
    Company_No.Keys(EnvParams.Opco);
    
    var Company_No = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McTextWidget2;
    Sys.HighlightObject(Company_No);
    Company_No.Click();
    Company_No.Keys(EnvParams.Opco);
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var ShowLines = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite3.McPlainCheckboxView.Button;
    Sys.HighlightObject(ShowLines);
    ShowLines.Click();
    aqUtils.Delay(3000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var Save = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(Save);
    Save.Click();
    aqUtils.Delay(3000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    
  }


}


function Change_Allowance_Period(){ 
  
var Employee_Early_Allowance = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(Employee_Early_Allowance);
Employee_Early_Allowance.Click();

    aqUtils.Delay(3000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
var Grid = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(Grid);

for(var i=0;i<7;i++){ 
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
}
var Total_Allowance_Period = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Total_Allowance_Period.setText(Total_Allowance);

aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    
var Save = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(Save);
Save.Click();

aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
}




function Submit_and_Approve_Allowance(){ 
  

var Submit_Adjustment = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(Submit_Adjustment);
Submit_Adjustment.Click();

aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Administration - Absence Transfer").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Administration - Absence Transfer").OleValue.toString().trim(), 2000);
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

}

aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Approve_Adjustment = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
Sys.HighlightObject(Approve_Adjustment);
Approve_Adjustment.Click();

aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Administration - Absence Transfer").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Absence Administration - Absence Transfer").OleValue.toString().trim(), 2000);
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

}

aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid

  Total_Allowance = parseFloat(Total_Allowance);
  Total_Allowance = Total_Allowance.toFixed(2);
  var flag = true;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(5).OleValue.toString().trim()==Total_Allowance){ 
      ValidationUtils.verify(true,true, "Given Allowance is Updated in Approved Allowance, Cur Period Column in maconomy");
      flag = false;
      break;
    }
    }
    
    if(flag)
    ValidationUtils.verify(true,false, "Given Allowance is NOT Updated in Approved Allowance, Cur Period Column in maconomy");
    
    
}