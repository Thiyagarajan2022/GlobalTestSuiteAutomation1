//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "JobCreation";
var Language = "";
  Indicator.Show();
  
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile;
 
//Main Function
function createJob() {
TextUtils.writeLog("Job Creation Started"); 
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
sheetName = "JobCreation";


ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile ="";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

getDetails();
goToJobMenuItem();   
createAJob();   
GoToJob();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();


}


//getting data from datasheet
function getDetails(){

ExcelUtils.setExcelName(workBook, sheetName, true);
comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((comapany==null)||(comapany=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");
}
Job_group = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)
if((Job_group==null)||(Job_group=="")){ 
ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");
}
Job_Type = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
if((Job_Type==null)||(Job_Type=="")){ 
ValidationUtils.verify(false,true,"Job Type Number is Needed to Create a Job");
}
department = ExcelUtils.getRowDatas("Department",EnvParams.Opco)
if((department==null)||(department=="")){ 
ValidationUtils.verify(false,true,"Department Number is Needed to Create a Job");
}
buss_unit = ExcelUtils.getRowDatas("BusinessUnit",EnvParams.Opco)
if((buss_unit==null)||(buss_unit=="")){ 
ValidationUtils.verify(false,true,"BusinessUnit Number is Needed to Create a Job");
}
TemplateNo = ExcelUtils.getRowDatas("Template",EnvParams.Opco)
if((TemplateNo==null)||(TemplateNo=="")){ 
ValidationUtils.verify(false,true,"Template Number is Needed to Create a Job");
}
ExcelUtils.setExcelName(workBook, "Data Management", true);
Product = ReadExcelSheet("Global Product Number",EnvParams.Opco,"Data Management");
if((Product=="")||(Product==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Product = ExcelUtils.getRowDatas("Product",EnvParams.Opco)
}
if((Product==null)||(Product=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Create a Job");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
Job_name= ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
if((Job_name==null)||(Job_name=="")){ 
ValidationUtils.verify(false,true,"Job Name is Needed to Create a Job");
}
Dlang= ExcelUtils.getRowDatas("Language",EnvParams.Opco)

BFC= ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)

pTerm= ExcelUtils.getRowDatas("Payment Terms",EnvParams.Opco)
ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

}




function JobAddress(){ 
//Checking Labels in Job Create Wizard

Sys.HighlightObject(Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", ""))
var companyName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McTextWidget", "", 1,60000).getText().OleValue.toString().trim();
//Log.Message(companyName);
//Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", companyName).OleValue.toString().trim())
if(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", companyName).OleValue.toString().trim()!="Company")
ValidationUtils.verify(false,true,"Company field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Company field is available in Maconomy for Job Creation");

var job = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
Log.Message(job);
Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,"English", job).OleValue.toString().trim())
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







// Providing Details for Job Creation
function createAJob() {
ReportUtils.logStep("INFO", "Enter Job Details");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  

  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var all_job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(all_job);
all_job.HoverMouse();
ReportUtils.logStep_Screenshot("");
  all_job.Click();

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    
  var newJobBtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  WorkspaceUtils.waitForObj(newJobBtn);
  newJobBtn.HoverMouse();
ReportUtils.logStep_Screenshot("");
  newJobBtn.Click();
TextUtils.writeLog("New Job is clicked");
aqUtils.Delay(3000, "Checking Labels");

var cancelJob = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
WorkspaceUtils.waitForObj(cancelJob)

  JobAddress();
  
//----------Entering Company Number-------------
ReportUtils.logStep_Screenshot("");
var companyName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,60000);
if(comapany!=""){
companyName.Click();
var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(companyName,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),comapany,ExlArray,"Company Number");
}else{ 
  ValidationUtils.verify(false,true,"Company is Needed to Create Job");
}
  
//----------Entering Job Group-------------

  var job = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);

if(Job_group!=""){
job.Click();
aqUtils.Delay(3000, "Waiting for ScrolledComposite");
var list = "";
try{
list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
}
catch(e){ 
 job.Click(); 
 aqUtils.Delay(3000, "Waiting for ScrolledComposite");
 list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
}
var Add_Visible2 = true;
while(Add_Visible2){
if(list.isEnabled()){
Add_Visible2 = false;
var dropstatus = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Job_group.toString().trim()){ 
          list.Keys("[Enter]");
          aqUtils.Delay(1000, Indicator.Text);
          dropstatus = true;
          ValidationUtils.verify(true,true,"Job Group is listed and selected in Maconomy");
          break;
        }else{ 
          list.Keys("[Down]");
        }
          
      }else{ 
        list.Keys("[Down]");
      }
    }
    if(!dropstatus)
    ValidationUtils.verify(false,true,"Job Group is not listed in Maconomy");
}
}
}
else{ 
    ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");
  }
  


//----------Entering Job Type-------------
   
  var JobType = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Job_Type!=""){
JobType.Click();
  var JG = "";
  if(Job_group.toString().trim()=="Client Billable")
  JG = "ValidateJobtype_CB(Client Billable)";
  if(Job_group.toString().trim()=="Client Non-Billable")
  JG = "ValidateJobtype_CNB(Client NonBillable)";
  if(Job_group.toString().trim()=="Internal")
  JG = "ValidateJobtype_IN(Internal)";
var ExlArray = []; 
ExlArray = getExcelData(JG,EnvParams.Opco);
if(ExlArray.length>0){ 
  
}
else
ValidationUtils.verify(false,true,"Selected Job Group doesn't have any Job Type in Opco's");
Job_Type = WorkspaceUtils.config_with_Maconomy_Validation(JobType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Type").OleValue.toString().trim(),Job_Type,ExlArray,"Job Type");
}else{ 
  ValidationUtils.verify(false,true,"JobType is Needed to Create Job");
}

//----------Entering Department-------------   
    
var Depart = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(department!=""){
Depart.Click();
var ExlArray = getExcelData("Validate_Department",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(Depart,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Department").OleValue.toString().trim(),department,ExlArray,"Department Number");
}else{ 
  ValidationUtils.verify(false,true,"Department is Needed to Create Job");
}
 
//----------Entering BusinessUnit-------------   
    
  var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(buss_unit!=""){
BussUnit.Click();
var ExlArray = getExcelData("Validate BusinessUnit",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(BussUnit,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Business Unit").OleValue.toString().trim(),buss_unit,ExlArray,"Business Unit Number");
}else{ 
  ValidationUtils.verify(false,true,"Business Unit is Needed to Create Job");
}

//----------Entering Template Number-------------    
  var template = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(TemplateNo!=""){
template.Click();
var ExlArray = getExcelData("Validate Template",EnvParams.Opco)
WorkspaceUtils.Config_with_Maconomy_templateValidation(template,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),TemplateNo,ExlArray,Job_Type,comapany,Job_group,"Template Number");
}else{ 
  ValidationUtils.verify(false,true,"Template is Needed to Create Job");
}
  
    
//----------Entering Product Number-------------    
    
    
  var prdNumber = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Product!=""){
prdNumber.Click();
WorkspaceUtils.SearchByValuePicker_Col_2(prdNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Product Result").OleValue.toString().trim(),Product,"Product Number");
}else{ 
  ValidationUtils.verify(false,true,"Product Number is Needed to Create Job");
}
    
//----------Entering Job Name-------------   
    
    
  var jobName = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
  jobName.setText(Job_name.toString().trim()+" "+STIME);
  if((jobName.getText().OleValue.toString().trim()==null)||(jobName.getText().OleValue.toString().trim()==""))
  ValidationUtils.verify(false,true,"Job Name can't able to enter in Maconomy");
  else
  ValidationUtils.verify(true,true,"Job Name is enter in Maconomy");

  
//----------Entering Project Manager-------------
  
var ProjectManger = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);  
 if(Project_manager!=""){
 if(ProjectManger.getText()!=Project_manager.toString().trim()){
 ProjectManger.Click();
 WorkspaceUtils.SearchByValues_Wiz2_Col_2(ProjectManger,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Project Manager");
 }
 }
 
//----------Clicking Create Button or Cancel Button-------------
var btnCreate = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());    
if(btnCreate.isEnabled()){

  Sys.HighlightObject(btnCreate)
  btnCreate.HoverMouse();
ReportUtils.logStep_Screenshot("");
  btnCreate.Click();
TextUtils.writeLog("Job is CREATED");
ValidationUtils.verify(true,true,"Job is CREATED");
ReportUtils.logStep("INFO", Job_name+" "+STIME +" : is Created");
TextUtils.writeLog("Job Name :"+Job_name+" "+STIME);
aqUtils.Delay(5000, "Waiting to Check any pop-up for credit limit");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
Log.Message("Language :"+Language);
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

 aqUtils.Delay(10000, "Reporting in HTML about Notification");  
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

 aqUtils.Delay(10000, "Reporting in HTML about Notification");  
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

}
}
else{
  
Log.Message("Language :"+Language);
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

//var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 1).SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//ReportUtils.logStep("INFO",lab)
//TextUtils.writeLog(lab);
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 1).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//Ok.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//Ok.Click();

 aqUtils.Delay(10000, "Reporting in HTML about Notification");  
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim()){ 
//var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim()).SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//ReportUtils.logStep("INFO",lab)
//TextUtils.writeLog(lab);
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//Ok.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//Ok.Click();
//} 
 aqUtils.Delay(10000, "Reporting in HTML about Notification"); 

  var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim(), 2000);
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

//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim()){ 
//var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim()).SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//ReportUtils.logStep("INFO",lab)
//TextUtils.writeLog(lab);
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//Ok.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//Ok.Click();
//} 

}
  

}

  
}
else{ 
var cancel = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
Sys.HighlightObject(cancel)
cancel.HoverMouse();
ReportUtils.logStep_Screenshot("");
cancel.Click();
ValidationUtils.verify(true,false,"Job is not Created");
ReportUtils.logStep("ERROR", "Job is not Created");
}
    
  aqUtils.Delay(4000, Indicator.Text);
}

//Validating Created Job is available
function GoToJob() {
  var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  WorkspaceUtils.waitForObj(table);
  var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  WorkspaceUtils.waitForObj(companyFilter);
  companyFilter.forceFocus();
  companyFilter.setVisible(true);
  companyFilter.ClickM();
  table.Child(0).setText("^a[BS]");
  table.Child(0).setText(comapany);
  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);

  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);

  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  WorkspaceUtils.waitForObj(job);
  job.Click();

  job.setText(Job_name+" "+STIME);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(2000, "Reading Table Data in Job List");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(Job_name+" "+STIME)){ 
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Job Number",EnvParams.Opco,"Data Management",table.getItem(v).getText_2(2).OleValue.toString().trim())
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  TextUtils.writeLog("Created Job is available in system");
  TextUtils.writeLog("Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  ReportUtils.logStep_Screenshot("");
  closeFilter.Click();
  aqUtils.Delay(4000, "Created Job Details is loading");
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  WorkspaceUtils.waitForObj(info);
info.HoverMouse();
ReportUtils.logStep_Screenshot("");
  info.Click();
  count=false;
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var Templete_Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Blanket_invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  var estimation = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Amount_Registrations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  
//  if(EnvParams.Country.toUpperCase()=="INDIA")
  var Invocing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 9).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
//  else
//  var Invocing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", ""); 
  
  var TimeReg = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  
  Sys.HighlightObject(Blanket_invoice)
  WorkspaceUtils.waitForObj(Blanket_invoice);
  WorkspaceUtils.waitForObj(estimation);
  WorkspaceUtils.waitForObj(Amount_Registrations);
  WorkspaceUtils.waitForObj(Invocing);
  WorkspaceUtils.waitForObj(TimeReg);
  checkmark = false;
  
//----------De-Select CheckBox-------------
  if(Templete_Job.getSelection()){ 
  Templete_Job.HoverMouse();
ReportUtils.logStep_Screenshot("");
  Templete_Job.Click();
  ReportUtils.logStep("INFO", "Templete_Job is UnChecked");
    Log.Message("Templete_Job is UnChecked")
    checkmark = true;
  }
  if(Blanket_invoice.getSelection()){ 
Blanket_invoice.HoverMouse();
ReportUtils.logStep_Screenshot("");
  Blanket_invoice.Click();
  ReportUtils.logStep("INFO", "Blanket_invoice is UnChecked");
//    Log.Message("Blanket_invoice is UnChecked")
    checkmark = true;
  }
  if(Amount_Registrations.getSelection()){ 
  Amount_Registrations.HoverMouse();
ReportUtils.logStep_Screenshot("");
  Amount_Registrations.Click();
  ReportUtils.logStep("INFO", "Amount_Registrations is UnChecked");
//    Log.Message("Amount_Registrations is UnChecked")
    checkmark = true;
  }
  if(Invocing.getSelection()){ 
  Invocing.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  Invocing.Click();
  ReportUtils.logStep("INFO", "Invocing is UnChecked");
//    Log.Message("Invocing is UnChecked")
    checkmark = true;
  }
  if(TimeReg.getSelection()){ 
  TimeReg.HoverMouse();
ReportUtils.logStep_Screenshot("");
  TimeReg.Click();
  ReportUtils.logStep("INFO", "TimeRegistration is UnChecked");
//    Log.Message("TimeRegistration is UnChecked")
    checkmark = true;
  }
  if(estimation.getSelection()){ 
  estimation.HoverMouse();
ReportUtils.logStep_Screenshot("");
  estimation.Click();
  ReportUtils.logStep("INFO", "Estimating is UnChecked");
//    Log.Message("Estimating is UnChecked")
    checkmark = true;
  }
  
  
  
  var JobLang = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2);
  var counterBFC = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  var paymentTerm = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2);
  var JobType = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  var Depart = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  var ProjectManger  =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);

  if((Dlang!="")&&(Dlang!=null)){
  if(JobLang.getText()!=Dlang){ 
  JobLang.Keys(" ")
  JobLang.Click();
  WorkspaceUtils.DropDownList(Dlang,"Language")
  }
  }
  if((BFC!="")&&(BFC!=null)){
  if(counterBFC.getText()!=BFC){ 
  counterBFC.Click();
  WorkspaceUtils.SearchByValue(counterBFC,"Counter Party BFC",BFC,"Counter Party BFC");
  }
  }
  if((pTerm!="")&&(pTerm!=null)){
  if(paymentTerm.getText()!=pTerm){ 
  paymentTerm.Keys(" ")
  paymentTerm.Click();
  WorkspaceUtils.DropDownList(pTerm,"Payment Term") 
  }
  }
  ExcelUtils.setExcelName(workBook, sheetName, true);
  Job_Type = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
  if(JobType.getText()!=Job_Type){ 
  JobType.Click();
  WorkspaceUtils.SearchByValue(JobType,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Type").OleValue.toString().trim(),Job_Type,"Job Type");
  }
  if(Depart.getText()!=department){ 
  Depart.Click();
  WorkspaceUtils.SearchByValue(Depart,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Department").OleValue.toString().trim(),department,"Department Number"); 
  }
  if(BussUnit.getText()!=buss_unit){ 
  BussUnit.Click();
  WorkspaceUtils.SearchByValue(BussUnit,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Business Unit").OleValue.toString().trim(),buss_unit,"Business Unit Number"); 
  }
  if(ProjectManger.getText()!=Project_manager){ 
  ProjectManger.Click();
  WorkspaceUtils.SearchByValues_Wiz2_Col_2(ProjectManger,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Project Manager");
  }
  

  if(checkmark){ 
var save_change = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.Save;
WorkspaceUtils.waitForObj(save_change);
save_change.HoverMouse();
ReportUtils.logStep_Screenshot("");
save_change.Click();
ValidationUtils.verify(true,true,"Changes is Saved");
TextUtils.writeLog("Changes is Saved");
aqUtils.Delay(5000, "Saving changes  in Jobs");
  }
  ReportUtils.logStep("INFO", "Job is Saved");
  TextUtils.writeLog("Job is Saved");
  var filter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
  filter.Click();
}
}

//Go To Job from Menu
function goToJobMenuItem(){
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}

function getExcelData(rowidentifier,column) { 
var temp = ""

var excelData = [];

ExcelUtils.setExcelName(workBook, sheetName, true);
temp = ExcelUtils.getRowDatas(rowidentifier,column);
 
 if(temp.indexOf(",")!=-1){ 
   excelData =  temp.split(",");
  }else if(temp.length>0){ 
      excelData[0] = temp;
     }

 return excelData;
}

function getExcelData_Company(rowidentifier,column) { 
var excelData =[];  
var temp ="";
ExcelUtils.setExcelName(workBook, sheetName, true);
temp = ExcelUtils.getRowDatas(rowidentifier,column);
 
if(temp.indexOf("*")!=-1){
var excelData =  temp.split("*");
}else if(temp.length>0){ 
excelData[0] = temp;
}
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



