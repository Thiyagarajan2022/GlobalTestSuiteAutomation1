//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

//Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ImportBudgetModel";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var jobtypevalue,jobdepartmenvalue,jobBusinessUnitvalue ="";
var totalAmount = "";
var journalNo  ="";

function ImportBudgetModel(){ 
TextUtils.writeLog("Create Import Budget Model"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);
 
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ImportBudgetModel";
STIME = "";
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Import Budget Model started::"+STIME);
jobtypevalue,jobdepartmenvalue,jobBusinessUnitvalue,journalNo,totalAmount ="";
getDetails();
gotoMenu();
gotoJournals();
addBudgetline();
WorkspaceUtils.closeAllWorkspaces();
}


function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();


if(ImageRepository.ImportBudgetModel.FinanceBudget1.Exists()){
ImageRepository.ImportBudgetModel.FinanceBudget1.Click();
}
else if(ImageRepository.ImportBudgetModel.FinanceBudget.Exists()){
ImageRepository.ImportBudgetModel.FinanceBudget.Click();
}
else
{
  ImageRepository.ImportBudgetModel.FinanceBudget2.Click();
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
Client_Managt.ClickItem("|Budget");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Budget");
}
}


}


function getDetails(){ 
  
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobtypevalue = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)

  if((jobtypevalue==null)||(jobtypevalue=="")){ 
  ValidationUtils.verify(false,true,"Job_group is Needed");
  }
    Log.Message("jobtypevalue"+jobtypevalue)
    
     
  jobdepartmenvalue = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
  if((jobdepartmenvalue==null)||(jobdepartmenvalue=="")){ 
  ValidationUtils.verify(false,true,"Job_Type is Needed");
  }
  
  Log.Message("jobdepartmenvalue"+jobdepartmenvalue)
  
  jobBusinessUnitvalue = ExcelUtils.getRowDatas("Job_Department",EnvParams.Opco)
if((jobBusinessUnitvalue==null)||(jobBusinessUnitvalue=="")){ 
ValidationUtils.verify(false,true,"Job_Department is Needed");

}
Log.Message("jobBusinessUnitvalue"+jobBusinessUnitvalue)


  totalAmount = ExcelUtils.getRowDatas("Total Amount",EnvParams.Opco)
if((totalAmount==null)||(totalAmount=="")){ 
ValidationUtils.verify(false,true,"totalAmount is Needed");

}
  

}

function gotoJournals(){ 
  
var  JournalsTab = Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(JournalsTab);
JournalsTab.Click();

var ImportButton =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.ImportButton;
WorkspaceUtils.waitForObj(ImportButton);
ImportButton.Click();

var ImportBudgetInformation =Aliases.Maconomy.ImportBudgetInfoPopup;
WorkspaceUtils.waitForObj(ImportBudgetInformation);
  var internalNames = Aliases.Maconomy.ImportBudgetInfoPopup.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McPlainCheckboxView.InternalNames;
  var progressBar = Aliases.Maconomy.ImportBudgetInfoPopup.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McPlainCheckboxView.ProgressBar
  var logging = Aliases.Maconomy.ImportBudgetInfoPopup.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McPlainCheckboxView.Button;

  //----------Select CheckBox-------------
  
  aqUtils.Delay(3000, "Waiting for pop-up");
  if(internalNames.getSelection()==false){ 
  internalNames.HoverMouse();
ReportUtils.logStep_Screenshot("");
  internalNames.Click();
  ReportUtils.logStep("INFO", "internalNames is UnChecked");
    Log.Message("internalNames is UnChecked")
    checkmark = true;
  }
  
   if(progressBar.getSelection()==false){ 
  progressBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
  progressBar.Click();
  ReportUtils.logStep("INFO", "progressBar is UnChecked");
    Log.Message("progressBar is UnChecked")
    checkmark = true;
  }
  
   if(logging.getSelection()==false){ 
  logging.HoverMouse();
ReportUtils.logStep_Screenshot("");
  logging.Click();
  ReportUtils.logStep("INFO", "logging is UnChecked");
    Log.Message("logging is UnChecked")
    checkmark = true;
  }
  
  var importButton =Aliases.Maconomy.ImportBudgetInfoPopup.Composite.Composite.Composite2.Composite.ImportButton;
  WorkspaceUtils.waitForObj(importButton);
  importButton.Click();
   
  var sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
  var sFileName = "Budget Model Template.txt";
 if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
  
    aqUtils.Delay(4000, "Waiting to Open file");;
    var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
    WorkspaceUtils.waitForObj(dicratory);
    dicratory.Keys(sFolder+sFileName);
    var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
    Sys.HighlightObject(opendoc);
    WorkspaceUtils.waitForObj(opendoc);
    opendoc.HoverMouse();
    ReportUtils.logStep_Screenshot();
    
    opendoc.Click();
    aqUtils.Delay(2000, "Document Attached");
    
    var save = Sys.Process("Maconomy").Window("#32770", "Save file", 1).Window("Button", "&Save", 1);
    Sys.HighlightObject(save);
    save.Click();
    
    aqUtils.Delay(2000, "Waiting For Completion");
    
    if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Save file")    
    {
    var button = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    var label = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Label", "*").WndCaption;
    Log.Message(label );
    button.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    button.Click();
    Delay(2000);
    }
       
    if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Budget - Import Budget Information")    
    {
    var button = Sys.Process("Maconomy").SWTObject("Shell", "Budget - Import Budget Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    var label = Sys.Process("Maconomy").SWTObject("Shell", "Budget - Import Budget Information").SWTObject("Label", "*").WndCaption;
    Log.Message(label );
    button.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    button.Click();
    Delay(2000);
    } 
      
    if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Save file")    
    {
    var button = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    var label = Sys.Process("Maconomy").SWTObject("Shell", "Save file").SWTObject("Label", "*").WndCaption;
    Log.Message(label );
    button.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    button.Click();
    Delay(2000);
    }

}


function addBudgetline()
{
  
var refresh =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.RefreshButton;
  refresh.Click();

 aqUtils.Delay(3000, "Wait for Refresh");
var table = Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JournalTable;

var flag;

  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==(totalAmount)){ 
      
    journalNo = table.getItem(v).getText_2(0).OleValue.toString().trim();
    Log.Message(journalNo)
    Log.Message("true")
    //  ExcelUtils.WriteExcelSheet("Job Number",EnvParams.Opco,"Data Management",table.getItem(v).getText_2(2).OleValue.toString().trim())
    flag=true;
    table.Keys("[PageUp][PageUp]");
    break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  table.Keys("[PageUp][PageUp]");
  
  ValidationUtils.verify(flag,true,"Budget Model is imported");
  ValidationUtils.verify(true,true,"Journal Number :"+table.getItem(v).getText_2(0).OleValue.toString().trim());
  TextUtils.writeLog("Created Job is available in system")
 
  var firsrcellBudgetModelTable = Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JournalTable.FirstceellTable;

  var firsrcellBudgetModelTableUnselected =
  firsrcellBudgetModelTable.Keys(journalNo);
  aqUtils.Delay(3000, "Wait for Table Load");
  
  var closefilter =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.Closefilter;
  
  closefilter.Click(); 
  
  var addBudgetlines = Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  addBudgetlines.Click();

  var firstcellsecondline =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SecondCellFirstLine;

  firstcellsecondline.Click();

  firstcellsecondline.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]")

  var jobType = Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "", 2);
  if(jobtypevalue!=""){
  jobType.Click();
  WorkspaceUtils.SearchByValue(jobType,"Local Specification 1",jobtypevalue,"Job Type");
  }
  
  jobType.Keys("[Tab]")
  
  var jobDepartment =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.JobType;
  if(jobdepartmenvalue!=""){
  jobDepartment.Click();
  WorkspaceUtils.SearchByValue(jobDepartment,"Local Specification 2",jobdepartmenvalue,"Job Department");
  }
  
   jobDepartment.Keys("[Tab][Tab]")  
   var jobbusinessUnit =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.JobType;

  if(jobBusinessUnitvalue!=""){
  jobbusinessUnit.Click();
  WorkspaceUtils.SearchByValue(jobbusinessUnit,"Local Specification 4",jobBusinessUnitvalue,"Job Business Unit");
  }
    
  var approve =Aliases.Maconomy.ImportBudgetModel.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.approveButton 
  approve.Click();
  
     aqUtils.Delay(2000, "Wait for Completion");
  
         if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Budget - Lines")    
{
var button = Sys.Process("Maconomy").SWTObject("Shell", "Budget - Lines").SWTObject("Composite", "", 2).SWTObject("Button", "Yes");
 var label = Sys.Process("Maconomy").SWTObject("Shell", "Budget - Lines").SWTObject("Label", "*").WndCaption;
      Log.Message(label );
       button.HoverMouse();
   //  ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(2000);
     }
}


function test()
{
  
         if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Budget - Lines")    
{
var button = Sys.Process("Maconomy").SWTObject("Shell", "Budget - Lines").SWTObject("Composite", "", 2).SWTObject("Button", "Yes");
 var label = Sys.Process("Maconomy").SWTObject("Shell", "Budget - Lines").SWTObject("Label", "*").WndCaption;
      Log.Message(label );
       button.HoverMouse();
   //  ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(2000);
     }
}