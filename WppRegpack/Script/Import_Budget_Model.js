//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

/**
 * This script to Import Budget Template for General Ledger
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :03/02/2021
*/

//Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ImportBudgetModel";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var JobType,Dept,Buss_Unit,Total_Amount,Journal_No,Control_9,Control_10,Control_11,Control_12 = "";





//Main Function
function Import_Budget_Template() {
  
TextUtils.writeLog("Import Budget Template Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Import Budget Template script
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
sheetName = "ImportBudgetModel";
STIME = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Import Budget Model started::"+STIME);
JobType,Dept,Buss_Unit,Total_Amount,Journal_No,Control_9,Control_10,Control_11,Control_12 = "";

getDetails();
goTo_Finance_Budget_Item();
Import_Files();
validatingJournal_In_Maconomy();


}




//Getting Data from Datasheets
function getDetails(){ 
  
JobType,Dept,Buss_Unit,Total_Amount,Journal_No,Control_9,Control_10,Control_11,Control_12 = "";

  ExcelUtils.setExcelName(workBook, "ImportBudgetModel", true);
  JobType = ExcelUtils.getRowDatas("Job Type",EnvParams.Opco);
  Dept = ExcelUtils.getRowDatas("Department",EnvParams.Opco);
  Buss_Unit = ExcelUtils.getRowDatas("Bussiness Unit",EnvParams.Opco);
  Total_Amount = ExcelUtils.getRowDatas("Total Amount",EnvParams.Opco);
  Control_9 = ExcelUtils.getRowDatas("Control 9",EnvParams.Opco);
  Control_10 = ExcelUtils.getRowDatas("Control 10",EnvParams.Opco);
  Control_11 = ExcelUtils.getRowDatas("Control 11",EnvParams.Opco);
  Control_12 = ExcelUtils.getRowDatas("Control 12",EnvParams.Opco);
  
  
  if((JobType=="")||(JobType==null)){
ValidationUtils.verify(true,false,"Job Type is needed to Import Budget Model")
  }  

    if((Dept=="")||(Dept==null)){
ValidationUtils.verify(true,false,"Department is needed to Import Budget Model")
  }  
  
    if((Buss_Unit=="")||(Buss_Unit==null)){
ValidationUtils.verify(true,false,"Business Unit is needed to Import Budget Model")
  }  
  
    if((Total_Amount=="")||(Total_Amount==null)){
ValidationUtils.verify(true,false,"Total Amount is needed to Import Budget Model")
  }  

      if((Control_9=="")||(Control_9==null)){
ValidationUtils.verify(true,false,"Control 9 is needed to Import Budget Model")
  }  
      if((Control_10=="")||(Control_10==null)){
ValidationUtils.verify(true,false,"Control 10 is needed to Import Budget Model")
  }  
      if((Control_11=="")||(Control_11==null)){
ValidationUtils.verify(true,false,"Control 11 is needed to Import Budget Model")
  }  
      if((Control_12=="")||(Control_12==null)){
ValidationUtils.verify(true,false,"Control 12 is needed to Import Budget Model")
  }  
}


/**
  *  This function Navigates to Budget screen from Finance Budget workspace
  */
function goTo_Finance_Budget_Item(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();

if(ImageRepository.ImageSet01.Finance_Budget.Exists()){
ImageRepository.ImageSet01.Finance_Budget.Click();
}
else if(ImageRepository.ImageSet01.Finance_Budget_1.Exists()){
ImageRepository.ImageSet01.Finance_Budget_1.Click();
}
else
{
  ImageRepository.ImageSet01.Finance_Budget_2.Click();
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Budget from Finance Budget Menu");
TextUtils.writeLog("Entering into Budget from Finance Budget Menu");
}



function Import_Files(){ 
 aqUtils.Delay(5000, Indicator.Text); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Journal = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(Journal);
Journal.Click();


aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Import_icon = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(Import_icon);
Import_icon.Click();


aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var Internal_Names = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McPlainCheckboxView.Button;
if(!Internal_Names.getSelection())
Internal_Names.Click();
aqUtils.Delay(4000, Indicator.Text);


var Internal_Popup_Names = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McPlainCheckboxView.Button;
if(Internal_Popup_Names.getSelection())
Internal_Popup_Names.Click();
aqUtils.Delay(4000, Indicator.Text);

var Progress_Bar = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McPlainCheckboxView.Button;
if(!Progress_Bar.getSelection())
Progress_Bar.Click();
aqUtils.Delay(4000, Indicator.Text);


var Logging = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite4.McPlainCheckboxView.Button;
if(!Logging.getSelection())
Logging.Click();
aqUtils.Delay(4000, Indicator.Text);

var Echo = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite5.McPlainCheckboxView.Button;
if(Echo.getSelection())
Echo.Click();
aqUtils.Delay(4000, Indicator.Text);

var Help = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite6.McPlainCheckboxView.Button;
if(Help.getSelection())
Help.Click();
aqUtils.Delay(4000, Indicator.Text);

var Report_Error_Lines = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite9.McPlainCheckboxView.Button;
if(Report_Error_Lines.getSelection())
Report_Error_Lines.Click();
aqUtils.Delay(4000, Indicator.Text);

var Print_Log = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite7.McPlainCheckboxView.Button;
if(Print_Log.getSelection())
Print_Log.Click();
aqUtils.Delay(4000, Indicator.Text);

var Run_Mode = Aliases.Maconomy.Import_Budget_Information.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite8.McPopupPickerWidget;
Run_Mode.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import").OleValue.toString().trim());



aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }



var Import = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import Budget Information").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import").OleValue.toString().trim())
Import.Click();

aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var sFolder = Project.Path+"TestResource\\Import Budget Model\\";
var sFileName = EnvParams.Opco+"_Budget Model Template.txt";
//Finding File ia availble or NOT
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ }
else{
Log.Error("Could not create the folder " + sFolder);
}
}
  
aqUtils.Delay(4000, "Waiting to Open file");;
var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
WorkspaceUtils.waitForObj(dicratory);
dicratory.Keys(sFolder+sFileName);
aqUtils.Delay(3000, "Waiting to Open file");;
var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
Sys.HighlightObject(opendoc);
WorkspaceUtils.waitForObj(opendoc);
opendoc.HoverMouse();
ReportUtils.logStep_Screenshot();
    
opendoc.Click();
aqUtils.Delay(2000, "Document Attached");
  
var p = Sys.Process("Maconomy").Window("#32770", "Save file", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
 var Reportfile = Sys.Process("Maconomy").Window("#32770", "Save file", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
 Sys.HighlightObject(Reportfile);
 var FileName = sFolder+EnvParams.Opco+"_"+Reportfile.wText;
 Reportfile.Keys(FileName)
 aqUtils.Delay(2000, "Document is Saving");
saveAs.Click();
}
   
aqUtils.Delay(2000, "Waiting For Completion");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }    

// var p = Sys.Process("Maconomy");
// var w = p.FindChild("WndCaption", "Save file", 2000);
//  if (w.Exists)
//{ 
//  
//    var button = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Composite", "", 2).SWTObject("Button",  JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//    var label = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Label", "*").WndCaption;
//    Log.Message(label );
//    button.HoverMouse();
//    ReportUtils.logStep_Screenshot("");
//    button.Click();
//    Delay(2000);
//  }  

aqUtils.Delay(2000, "Waiting For Completion");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
    
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget - Import Budget Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Budget - Import Budget Information").OleValue.toString().trim(), 2000);
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
       
aqUtils.Delay(2000, "Waiting For Completion");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
      
// var p = Sys.Process("Maconomy");
// var w = p.FindChild("WndCaption", "Save file", 2000);
//  if (w.Exists)
//{ 
//    var button = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Composite", "", 2).SWTObject("Button",  JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//    var label = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Label", "*").WndCaption;
//    Log.Message(label );
//    button.HoverMouse();
//    ReportUtils.logStep_Screenshot("");
//    button.Click();
//    Delay(2000);
//    }

aqUtils.Delay(2000, "Waiting For Completion");
}



function validatingJournal_In_Maconomy(){ 
  
// aqUtils.Delay(5000, Indicator.Text); 
//if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
//
//var Journal = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.TabControl;
//Sys.HighlightObject(Journal);
//Journal.Click();


aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, "Waiting for Maconomy");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Refresh = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(Refresh);
Refresh.Click();


aqUtils.Delay(4000, "Waiting for Maconomy");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var FirstCell = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(FirstCell);
FirstCell.Click();
FirstCell.Keys("[Tab]");
aqUtils.Delay(2000, "Waiting for Maconomy");

var CompanyNo = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(CompanyNo);
CompanyNo.Click();
CompanyNo.Keys(EnvParams.Opco);

for(var i=0;i<12;i++){ 
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");
}
var Control_9_Col = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Control_9_Col);
Control_9_Col.Click();
Control_9_Col.Keys(Control_9);

  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");
  
var Control_10_Col = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Control_10_Col);
Control_10_Col.Click();
Control_10_Col.Keys(Control_10);

  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");

var Control_11_Col = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Control_11_Col);
Control_11_Col.Click();
Control_11_Col.Keys(Control_11);

  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");

var Control_12_Col = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Control_12_Col);
Control_12_Col.Click();
Control_12_Col.Keys(Control_12);

aqUtils.Delay(4000, "Waiting for Maconomy");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var flag = false;
  for(var v=0;v<table.getItemCount();v++){ 
    if((table.getItem(v).getText_2(1).OleValue.toString().trim()== EnvParams.Opco) &&
    (table.getItem(v).getText_2(13).OleValue.toString().trim()== Control_9) &&
    (table.getItem(v).getText_2(14).OleValue.toString().trim()== Control_10) &&
    (table.getItem(v).getText_2(15).OleValue.toString().trim()== Control_11) &&
    (table.getItem(v).getText_2(16).OleValue.toString().trim()== Control_12) ){ 
      JobID = table.getItem(v).getText_2(2).OleValue.toString().trim();
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

if(flag){
var closeFilter = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
Sys.HighlightObject(closeFilter);
closeFilter.Click();

aqUtils.Delay(4000, "Waiting for Maconomy");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Add_icon = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(Add_icon);
Add_icon.Click();

aqUtils.Delay(4000, "Waiting for Maconomy");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var FirstCell = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
Sys.HighlightObject(FirstCell);
FirstCell.Click();
FirstCell.Keys("[Tab]");

for(var i=0;i<16;i++){ 
  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");
}

//var Amount_Close = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//Sys.HighlightObject(Amount_Close);
//Amount_Close.Click();
//Amount_Close.setText(Total_Amount)


var Job_Type = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  Job_Type.Click();
  WorkspaceUtils.SearchByValue(Job_Type,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 1").OleValue.toString().trim(),JobType,"Job Type");
  
  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");
var Department = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  Department.Click();
  WorkspaceUtils.SearchByValue(Department,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 2").OleValue.toString().trim(),Dept,"Department");

  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");
  
  aqUtils.Delay(1000, "Waiting for Maconomy");
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, "Waiting for Maconomy");
  
var Business_unit = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  Business_unit.Click();
  WorkspaceUtils.SearchByValue(Business_unit,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 4").OleValue.toString().trim(),Buss_Unit,"Business Unit");

  aqUtils.Delay(3000, "Waiting for Maconomy");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
var Save = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(Save);
Save.Click();

  aqUtils.Delay(3000, "Waiting for Maconomy");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
var Approve = Aliases.Maconomy.Periodic_Balance.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(Approve);
Approve.Click();

  aqUtils.Delay(3000, "Waiting for Maconomy");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  


}



}