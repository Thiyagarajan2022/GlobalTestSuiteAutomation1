//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "JobClosure";
var Language = "";
  Indicator.Show();
  
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_number,Project_manager,OpCoFile;



//getting data from datasheet
function getDetails(){
//Log.Message("excelName :"+workBook);
//Log.Message("sheet :"+sheetName);
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(EnvParams.Opco)
comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((comapany==null)||(comapany=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");
}

 ExcelUtils.setExcelName(workBook, "Data Management", true);
  Job_number= ExcelUtils.getRowDatas("Job Number",EnvParams.Opco);
  if((Job_number==null)|| (Job_number==""))
  {
    ExcelUtils.setExcelName(workBook, sheetName, true);
    Job_number= ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
  }
  if((Job_number==null)||(Job_number=="")){ 
  ValidationUtils.verify(false,true,"Job_number is Needed to Create a Job");
}
ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}




function JobAddress(){ 
//Checking Labels in Job Create Wizard
//Delay(4000);
Sys.Process("Maconomy").Refresh();
var companyName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McTextWidget", "", 1,60000).getText().OleValue.toString().trim();
//Log.Message(companyName);
//Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", companyName).OleValue.toString().trim())
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", companyName).OleValue.toString().trim()!="Company")
ValidationUtils.verify(false,true,"Company field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Company field is available in Maconomy for Job Creation");
var job = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", job).OleValue.toString().trim()!="Job Group")
ValidationUtils.verify(false,true,"Job Group field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Job Group field is available in Maconomy for Job Creation");
var JobType = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", JobType).OleValue.toString().trim()!="Job Type")
ValidationUtils.verify(false,true,"Job Type field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Job Type field is available in Maconomy for Job Creation");
var Depart = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", Depart).OleValue.toString().trim()!="Department")
ValidationUtils.verify(false,true,"Department field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Department field is available in Maconomy for Job Creation");
var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", BussUnit).OleValue.toString().trim()!="Business Unit")
ValidationUtils.verify(false,true,"Business Unit field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Business Unit field is available in Maconomy for Job Creation");
var template = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", template).OleValue.toString().trim()!="Template")
ValidationUtils.verify(false,true,"Template field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Template field is available in Maconomy for Job Creation");
var prdNumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", prdNumber).OleValue.toString().trim()!="Product")
ValidationUtils.verify(false,true,"Product field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Product field is available in Maconomy for Job Creation");
var jobName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", jobName).OleValue.toString().trim()!="Job Name")
ValidationUtils.verify(false,true,"Job Name field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Job Name field is available in Maconomy for Job Creation");
var ProjectManger = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", ProjectManger).OleValue.toString().trim()!="Project Manager")
ValidationUtils.verify(false,true,"Project Manager field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Project Manager field is available in Maconomy for Job Creation");

}

//Validating Created Job is available
function GoToJob() {
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  var closeFilter = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  aqUtils.Delay(2000, Indicator.Text);
//  Delay(2000);

  var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  companyFilter.forceFocus();
  companyFilter.setVisible(true);
  companyFilter.ClickM();
  table.Child(0).setText("^a[BS]");
  table.Child(0).setText(comapany);
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x09);
//  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();

  job.setText(Job_number);
  aqUtils.Delay(7000, Indicator.Text);
//  Delay(7000);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==(Job_number)){ 
//  var ExlArray = getExcelData("Validate_Company",EnvParams.Opco)  
//   var name =  LogReport_name(ExlArray,comapany,Job_group);
//      var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";
////      Log.Message("Notepad :"+notepadPath)
//      TextUtils.writeDetails(notepadPath,"Job Number ",table.getItem(v).getText_2(2).OleValue.toString().trim());
     // ExcelUtils.setExcelName(workBook,"Data Management", true);
    //  ExcelUtils.WriteExcelSheet("Job Number",EnvParams.Opco,"Data Management",table.getItem(v).getText_2(2).OleValue.toString().trim())
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  ReportUtils.logStep_Screenshot("");
  closeFilter.Click();
  aqUtils.Delay(8000, Indicator.Text);
//  Delay(8000);
//    var Home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//    Home.Click();
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  }
  }


}

//Go To Job from Menu
function goToJobMenuItem(){

aqUtils.Delay(5000, Indicator.Text);
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

var mainlist = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var main;
for(var id=0;id<mainlist;id++){
main = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
if(main.Child(id).isVisible())
if(main.Child(id).ChildCount==1)
if(main.Child(id).Child(0).Name.indexOf("Composite")!=-1){

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot("");
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}
}
}

}
//Delay(5000); 
ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
 
}

//Main Function
function JobClosure() {
//Indicator.PushText("waiing for window to open");
//var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//
//menuBar.Click();
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//Restart.login(Project_manager);
//  
//}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "JobClosure";
Language = "";
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_number,Project_manager,OpCoFile ="";


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
ReportUtils.logStep("INFO", "Job Closure started::"+STIME);
getDetails();
goToJobMenuItem();      
GoToJob();
closeJob();
WorkspaceUtils.closeAllWorkspaces();
//aqTestCase.End();

}


function closeJob()
{
    Delay(8000);
  var jobClosingTab = Aliases.ObjectGroup.JobClosingTab;
  
//  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
jobClosingTab.HoverMouse();
  ReportUtils.logStep_Screenshot("");
jobClosingTab.Click();
  count=false;
  
  aqUtils.Delay(5000, Indicator.Text);
//  Delay(5000);
//var pendingActionsTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
  aqUtils.Delay(5000, Indicator.Text);
var closeJobButton = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;

 if(closeJobButton.isEnabled()){     

      closeJobButton.HoverMouse();

      ReportUtils.logStep_Screenshot();

     Sys.HighlightObject(closeJobButton);
  closeJobButton.Click();

  
          ValidationUtils.verify(true,true,"Job is Closed");
 ValidationUtils.verify(true,true,"Job Closure is successful");
 Log.Checkpoint("Job Closure is successful");
   

          }
        else{

          ValidationUtils.verify(false,true,"Close Job Button is invisible");

          ReportUtils.logStep("INFO","Job Closure not successful");
           Log.Checkpoint("Job Closure is not successful");

        }


   aqUtils.Delay(5000, Indicator.Text);

  ReportUtils.logStep("INFO", "Job is Closed");
//  var filter =
//  NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.JobClosureCloseFilter;
//  filter.Click();

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
//var array = "Validate_EmployeeCategories";
//var Opco = "1307"
var excelData = [];
//Log.Message("Execution completed,sending result to excel book , FileName:"+excelName+"sheetname:"+sheet);
  var app = Sys.OleObject("Excel.Application");
//  app.Visible = "True";
  var curArrayVals = [];  
//  Log.Message(workBook)
//  Log.Message(sheetName)
//  Log.Message(rowidentifier)
//  Log.Message(column)
  var book = app.Workbooks.Open(workBook);
  var sheet = book.Sheets.Item(sheetName);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;
//  Log.Message(columnCount);
//  Log.Message(rowCount);
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
//  Log.Message(sheet.Cells.Item(k, 1).Text.toString().trim())
  if(sheet.Cells.Item(k, 1).Text.toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
  row = k;
  rowStatus = true;
  }
  }
//  Log.Message(col)
//  Log.Message(row);
  if(rowStatus){ 
   temp = sheet.Cells.Item(row,  col).Text;
//   Log.Message(temp)
  }
  
  
// book.Save();
 app.Quit();
 
 
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
     

 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);

 return excelData;
}

function getExcelData_Company(rowidentifier,column) { 
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
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
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
