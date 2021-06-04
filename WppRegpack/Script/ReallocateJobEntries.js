//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ReallocateJobEntries";
var Language = "";
  Indicator.Show();
  
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var JobNo,Reallocation_Job_No,WorkCode,JournalNumber;
var Project_manager = "";


//getting data from datasheet
                                                                                                                                                                                      
                                                                                                                                                                                      
function getDetails(){

ExcelUtils.setExcelName(workBook, sheetName, true);



JobNo=ExcelUtils.getRowDatas("Re-allocate From Job Number",EnvParams.Opco)
if((JobNo==null)||(JobNo=="")){ 
  JobNo=ExcelUtils.getRowDatas("Job Serial Order From",EnvParams.Opco)
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  JobNo = ExcelUtils.getRowDatas("Job Number_"+JobNo,EnvParams.Opco)
}
if((JobNo==null)||(JobNo=="")){ 
ValidationUtils.verify(false,true,"Re-allocate From Job Number is Needed to Reallocating Job Entries");
}
Log.Message(JobNo)

ExcelUtils.setExcelName(workBook, sheetName, true);
Reallocation_Job_No=ExcelUtils.getRowDatas("Re-allocate To Job Number",EnvParams.Opco)
if((Reallocation_Job_No==null)||(Reallocation_Job_No=="")){ 
  Reallocation_Job_No=ExcelUtils.getRowDatas("Job Serial Order To",EnvParams.Opco)
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  Reallocation_Job_No = ExcelUtils.getRowDatas("Job Number_"+Reallocation_Job_No,EnvParams.Opco)
}
if((Reallocation_Job_No==null)||(Reallocation_Job_No=="")){ 
ValidationUtils.verify(false,true,"Re-allocate To Job Number is Needed to Reallocating Job Entries");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
WorkCode = ExcelUtils.getRowDatas("WorkCode",EnvParams.Opco)
if((WorkCode==null)||(WorkCode=="")){ 
ValidationUtils.verify(false,true,"WorkCode is Needed to Reallocating Job Entries");
}
Log.Message(WorkCode)

//JournalNumber = ExcelUtils.getRowDatas("JournalNumber",EnvParams.Opco)
//if((JournalNumber==null)||(JournalNumber=="")){ 
//ValidationUtils.verify(false,true,"JournalNumber is Needed to Reallocating Job Entries");
//}
//Log.Message(JournalNumber)


//ExcelUtils.setExcelName(workBook, "Server Details", true);
//Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}


//Go To Job from Menu
function goToJobMenuItem(){

//aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();


if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
}

//if(ImageRepository.ImageSet.Jobs1.Exists()){
//ImageRepository.ImageSet.Jobs1.Click();// GL
//}
//
//
//if(ImageRepository.ImageSet3.Jobs.Exists()){
// ImageRepository.ImageSet3.Jobs.Click();// GL
//}
//else if(ImageRepository.ImageSet.Job.Exists()){
//ImageRepository.ImageSet.Job.Click();
//}
//else{
////ImageRepository.ImageSet.Jobs1.Click();
// ImageRepository.ImageSet3.Jobs.Click();
//}

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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());


}

}



ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}


//Main Function
function ReallocateJobEntries() {
Indicator.PushText("waiting for window to open");
//aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)

menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
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
sheetName = "ReallocateJobEntries";
Language = "";
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
JobNo,Reallocation_Job_No,WorkCode,JournalNumber ="";


Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Log.Message(EnvParams.Opco)
Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Re-allocate Job Entries started::"+STIME);
getDetails();
goToJobMenuItem();      
GoToJob();
//WorkspaceUtils.closeAllWorkspaces();


}


function GoToJob() {
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading Data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading Data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading Data");

var JobNoTextBox = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.JobNo;

//Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.JobSearchField;

JobNoTextBox.setText(JobNo);
aqUtils.Delay(1000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading Data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading Data");
var table = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid

WorkspaceUtils.waitForObj(table);


var labels=Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.Label;
//Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.
//McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
      labels = labels.Child(i);
      break;
    }
  }

  WorkspaceUtils.waitForObj(labels);

  var i=0;
  while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=600)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}


  aqUtils.Delay(2000, "Reading Table Data in Job List");
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(JobNo)){ 

      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(0).OleValue.toString().trim());
aqUtils.Delay(1000);
  var closefilter = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter;
  //Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.Closefilter;
  closefilter.Click();

aqUtils.Delay(1000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading Data");



//  var closeFilter = Aliases.ObjectGroup.closeFilterJobAdminstration;
//  
//  //Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//
//  var table = Aliases.ObjectGroup.JobAdminTable;
//  
//  //Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//  aqUtils.Delay(2000, Indicator.Text);
////  Delay(2000);
//
//  var JobNoFilter = Aliases.ObjectGroup.JobAdminTable.SWTObject("McTextWidget", "");
//  
////  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
////  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
////  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
////  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
////  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
////  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//  JobNoFilter.forceFocus();
//  JobNoFilter.setVisible(true);
//  JobNoFilter.ClickM();
//  table.Child(0).setText("^a[BS]");
//  table.Child(0).setText(Job_No);
//  aqUtils.Delay(1000, Indicator.Text);
////  Delay(1000);
//  Sys.Desktop.KeyDown(0x09); // Press Ctrl
//  aqUtils.Delay(1000, Indicator.Text);
////  Delay(1000);
// // Sys.Desktop.KeyDown(0x09);
//  aqUtils.Delay(1000, Indicator.Text);
////  Delay(1000);
////  Sys.Desktop.KeyDown(0x09);
////  Sys.Desktop.KeyUp(0x09);
////  Sys.Desktop.KeyUp(0x09);
////  Sys.Desktop.KeyUp(0x09);
//  var job = Aliases.ObjectGroup.JobAdminTable.JobNameFilterJobAdministration
// // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
//  job.Click();
//
//  job.setText(Job_name);
//  aqUtils.Delay(7000, Indicator.Text);
////  Delay(7000);
//  var flag=false;
//  for(var v=0;v<table.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(Job_name)){ 
////  var ExlArray = getExcelData("Validate_Company",EnvParams.Opco)  
////   var name =  LogReport_name(ExlArray,comapany,Job_group);
////      var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";
//////      Log.Message("Notepad :"+notepadPath)
////      TextUtils.writeDetails(notepadPath,"Job Number ",table.getItem(v).getText_2(2).OleValue.toString().trim());
//     // ExcelUtils.setExcelName(workBook,"Data Management", true);
//    //  ExcelUtils.WriteExcelSheet("Job Number",EnvParams.Opco,"Data Management",table.getItem(v).getText_2(2).OleValue.toString().trim())
//      flag=true;
//      break;
//    }
//    else{ 
//      table.Keys("[Down]");
//    }
//  }
//
//  ValidationUtils.verify(flag,true,"Job Created is available in system");
//  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
//  if(flag){
//  ReportUtils.logStep("INFO", "Created Job is listed in table");
//  ReportUtils.logStep_Screenshot("");
//  closeFilter.Click();
//  aqUtils.Delay(8000, Indicator.Text);
////  Delay(8000);
////    var Home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
////    Home.Click();
//  if(count){
//  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
//  ref.Refresh();
//  }
//  }
  
  
  aqUtils.Delay(1000);
  
   var MarkedonlyCheckBox = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McPlainCheckboxView.Button
   //Aliases.ObjectGroup.MarkedOnlyJobAdministration
  // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  checkmark = false;
  if(MarkedonlyCheckBox.getSelection()){ 
  MarkedonlyCheckBox.HoverMouse();
ReportUtils.logStep_Screenshot("");
  MarkedonlyCheckBox.Click();
  ReportUtils.logStep("INFO", "MarkedonlyCheckBox is UnChecked");
    Log.Message("MarkedonlyCheckBox is UnChecked")
    checkmark = true;
  }
  

//var MarkForReallocation =Aliases.ObjectGroup.MarkForReallocation;
//MarkForReallocation.Click();


var jobReAllocateTo = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.ReallocateJobNo


if(Reallocation_Job_No!=""){
  jobReAllocateTo.Click();
  
  WorkspaceUtils.SearchByValues_all_Col_2(jobReAllocateTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),Reallocation_Job_No,"Job Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
 // WorkspaceUtils.SearchByValue(jobNoFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),JobNoFrom,"Job Number");
    }
 else{ 
    ValidationUtils.verify(false,true,"Reallocation_Job_No is Needed to Reallocate Job");
  }

  var saveButton =Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.MasterSave

saveButton.Click();
aqUtils.Delay(4000);
//jobReAllocateTo.Click();
//Delay(5000);
//jobReAllocateTo.setText(NewJob_No);
//Delay(5000);
//jobReAllocateTo.Keys("[Enter]");
                
//WorkspaceUtils.SearchByValueTableComp(jobReAllocateTo,"Job",comapany,"Company No.");


 
  var purchaseOrderTable =Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.JobReallocationTable;
  
//Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
  WorkspaceUtils.waitForObj(purchaseOrderTable);
 // purchaseOrderTable.Click();
  
    var flag=false;
 JournalNumber
   for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
  if(purchaseOrderTable.getItem(v).getText_2(11).OleValue.toString().trim()==WorkCode){ 
//    if(purchaseOrderTable.getItem(v).getText_2(4).OleValue.toString().trim()==(WorkCode)&&(purchaseOrderTable.getItem(v).getText_2(3).OleValue.toString().trim()==(JournalNumber))){ 

      flag=true;
//    purchaseOrderTable.Keys("[Tab]");
//    aqUtils.Delay(100);
//    purchaseOrderTable.Keys(" ");
//    aqUtils.Delay(1000);
//    purchaseOrderTable.Keys("[Tab]");
//    aqUtils.Delay(500);
//      purchaseOrderTable.Keys(" ");
//    aqUtils.Delay(1000);
    
var Job_Entries = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
Job_Entries.Click();
aqUtils.Delay(4000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var ReAllocate_CheckBox = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.JobReallocationTable.SWTObject("McPlainCheckboxView", "").SWTObject("Button", "");
if(!ReAllocate_CheckBox.getSelection()){ 
  ReAllocate_CheckBox.Click();
  aqUtils.Delay(4000,"Maconomy loading data");
}

purchaseOrderTable.Keys("[Tab]");
var Selected_CheckBox = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.JobReallocationTable.SWTObject("McPlainCheckboxView", "").SWTObject("Button", "");
if(!Selected_CheckBox.getSelection()){ 
  Selected_CheckBox.Click();
  aqUtils.Delay(4000,"Maconomy loading data");
}
JournalNumber = purchaseOrderTable.getItem(v).getText_2(6).OleValue.toString().trim()
  var savePOLine = Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SaveReallocation;
  
//Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SavePOLine

  savePOLine.Click();
  aqUtils.Delay(4000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
//    var MarkForAccrual =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.MarkForAccrual;
//  MarkForAccrual.Click();
    
  aqUtils.Delay(3000);
 
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Re-Allocated Entries From Job No",EnvParams.Opco,"Data Management",JobNo)
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Re-Allocated Entries To Job No",EnvParams.Opco,"Data Management",Reallocation_Job_No)
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Re-Allocated Journal No",EnvParams.Opco,"Data Management",JournalNumber)
  
  
  
      break;
      
    }
    else{ 
      purchaseOrderTable.Keys("[Down]");
    }
  }
  
   if(flag){
  ValidationUtils.verify(flag,true,"Journal with Work Code is available in system");
  ValidationUtils.verify(true,true,"Job Reallocation is Successful");
  }
  else{
     ValidationUtils.verify(false,true,"Journal with Work Code is not available in system");
  ValidationUtils.verify(false,true,"Job Reallocation is not Successful");
  }
      
var ApproveReallocation =Aliases.Maconomy.ReallocateJobEntries.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.ApproveReallocation
//Aliases.ObjectGroup.ApproveReallocation;
Sys.HighlightObject(ApproveReallocation);
ApproveReallocation.Click();
aqUtils.Delay(4000);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration - Job Reallocation by Job").OleValue.toString().trim())    
{
  Log.Message("Inside popup")
var button =Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration - Job Reallocation by Job").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());

//Sys.Process("Maconomy").SWTObject("Shell", "Administración de trabajo - Reasignación trabajo por trabajo").SWTObject("Composite", "", 2).SWTObject("Button", "Aceptar")

// Aliases.UserResubmitPopupOk
      var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration - Job Reallocation by Job").OleValue.toString().trim()).SWTObject("Label", "*").WndCaption;

      Log.Message(label );
       button.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      button.Click();
      Delay(5000);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  }   
  
  
}

function getExcel(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf(",")!=-1){
     var excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
}


function getExcelData(rowidentifier,column) { 
var temp = ""

var excelData = [];
//Log.Message(workBook+":")
//Log.Message(sheetName+":")
//Log.Message(rowidentifier+":")
//Log.Message(column+":")
ExcelUtils.setExcelName(workBook, sheetName, true);
temp = ExcelUtils.getRowDatas(rowidentifier,column);
//Log.Message(temp);
//temp = temp.OleValue.toString().trim();

/*
  var app = Sys.OleObject("Excel.Application");
  var curArrayVals = [];  
  var book = app.Workbooks.Open(workBook);
  var sheet = book.Sheets.Item(sheetName);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;

  var arrays={};
  var idx =0;
  var col =0;
  var row = 0;
  for(var k = 1; k<=columnCount;k++){
  if(sheet.Cells.Item(1, k).Text.toString().trim().toUpperCase()==column.toUpperCase()){
  col = k;
  }
  }
  var rowStatus = false;
  for(var k = 1; k<=rowCount;k++){
  if(sheet.Cells.Item(k, 1).Text.toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
  row = k;
  rowStatus = true;
  }
  }
  if(rowStatus){ 
   temp = sheet.Cells.Item(row,  col).Text;

  }
  
 app.Quit();
*/
 
 if(temp.indexOf(",")!=-1){ 
//       Log.Message(temp)
      excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     

// for(var i=0;i<excelData.length;i++)
// Log.Message(" :"+excelData[i]);

 return excelData;
}

function getExcelData_Company(rowidentifier,column) { 
var excelData =[];  
var temp ="";
ExcelUtils.setExcelName(workBook, sheetName, true);
temp = ExcelUtils.getRowDatas(rowidentifier,column);
//temp = temp.OleValue.toString().trim();

/*
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
DDT.CloseDriver(xlDriver.Name);
*/
     
     if(temp.indexOf("*")!=-1){
     var excelData =  temp.split("*");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
//     DDT.CloseDriver(xlDriver.Name);

// for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
}



function LogReport_name(ExcelData,value,JG){ 
var compStatus = "";
      for(var exl =0;exl<ExcelData.length;exl++){
        var splits = []; 
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
      if(splits[0]==value.toString().trim()){ 
        compStatus = ExcelData[exl]+"_"+JG;
        break;
      }
      }
Log.Message(compStatus);
return compStatus
}
