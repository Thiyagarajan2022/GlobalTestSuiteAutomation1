//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


/**
 * This script create job for Global Product
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :02/09/2021
*/

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "";
var Language = "";
Indicator.Show();

//Global Varibales
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile;
var Approve_Level = [];
var JobID,level,Foreign = ""; 
var flag=false;
var Jobnum = "";

//Main Function
function createJob(sheet,Jobnumber) {
  
TextUtils.writeLog("Job Creation Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Job Creation script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = sheet;
Jobnum = Jobnumber;

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile ="";
Approve_Level = [];
JobID,Foreign = "";
level = "";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

//try{

getDetails();
goToJobMenuItem();   
New_Job();
Create_Job_Wizard();   
Submiting_Job();

// Approving Job in Multi-Levels
for(var i=0;i<ApproveInfo.length;i++){
level = i;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);

// Restarting maconomy with Approver Logins
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);

//Refreshing To-Do's List to find Submitted Job
vID_Status = true;
todo(temp[3],temp[1],temp[2],i);

//Approving Created Job is every Levels
ApproveJob(temp[1],temp[2],i,temp[3]);
}

ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Job Number_"+Jobnum,EnvParams.Opco,"Data Management",JobID)
ExcelUtils.WriteExcelSheet("Job Template_"+Jobnum,EnvParams.Opco,"Data Management",TemplateNo)

//}
//  catch(err){
//    Log.Message(err);
//  }
  
// Close all opened workspace
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();


aqUtils.Delay(2000, Indicator.Text);

goToJobMenuItem(); 
ConvetToQuote();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(2000, Indicator.Text);

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


//Getting data from datasheet
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

Foreign= ExcelUtils.getRowDatas("Foreign Client",EnvParams.Opco)
if(Foreign.toUpperCase()=="YES"){ 
// Create Job Using Foreign Client
Log.Message("Foreign Client");
ExcelUtils.setExcelName(workBook, "Data Management", true);
Product = ReadExcelSheet("Foreign Global Product Number",EnvParams.Opco,"Data Management");
if((Product=="")||(Product==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Product = ExcelUtils.getRowDatas("Product",EnvParams.Opco)
}
if((Product==null)||(Product=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Create a Job");
}
  
}
else{
// Create Job Using Local Client
Log.Message("Local Client");
ExcelUtils.setExcelName(workBook, "Data Management", true);
Product = ReadExcelSheet("Global Product Number",EnvParams.Opco,"Data Management");
if((Product=="")||(Product==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Product = ExcelUtils.getRowDatas("Product",EnvParams.Opco)
}
if((Product==null)||(Product=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Create a Job");
}
}

ExcelUtils.setExcelName(workBook, sheetName, true);
Job_name= ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
if((Job_name==null)||(Job_name=="")){ 
ValidationUtils.verify(false,true,"Job Name is Needed to Create a Job");
}
Dlang= ExcelUtils.getRowDatas("Language",EnvParams.Opco)

BFC= ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)

pTerm= ExcelUtils.getRowDatas("Payment Terms",EnvParams.Opco)



ExcelUtils.setExcelName(workBook, sheetName, true);
Project_manager = ExcelUtils.getRowDatas("Project Manager",EnvParams.Opco)

}



/**
  *  This function Navigates to Jobs screen from Jobs workspace
  */
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


// Clicking New Job Button
function New_Job() {
ReportUtils.logStep("INFO", "Enter Job Details");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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

}




function Create_Job_Wizard(){
  

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
Log.Message(Project_manager)
 if((Project_manager!="")||(Project_manager!=null)){
 if(ProjectManger.getText()!=Project_manager.toString().trim()){
 ProjectManger.Click();
 WorkspaceUtils.SearchByValue(ProjectManger,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Project Manager");
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


//Clicking Pop-ups or Clicking Notification for multiple times
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



// Checking for 2nd Pop-ups
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


// Checking for 3rd Pop-ups
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


// Checking for 4th Pop-ups
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
  
// Checking for 1st Pop-ups
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


// Checking for 2nd Pop-ups
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


// Checking for 3rd Pop-ups
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


// Checking for 4th Pop-ups
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
function Submiting_Job() {
  
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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
  
  //Finding Created Job
  flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(Job_name+" "+STIME)){ 
      JobID = table.getItem(v).getText_2(2).OleValue.toString().trim();
      flag=true;
      table.Keys("[Down]");
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
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  //Validating Job and Invoice Currency
  aqUtils.Delay(2000, "Checking Job and Invoice Currency");
  var JobCurrency = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  var InvoiceCurrency = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  JobCurrency = JobCurrency.getText().OleValue.toString().trim();
  InvoiceCurrency = InvoiceCurrency.getText().OleValue.toString().trim();
  ReportUtils.logStep("Job Currency:"+JobCurrency)
  ReportUtils.logStep("Invoice Currency:"+InvoiceCurrency)
  Log.Message("JobCurrency :"+JobCurrency);
  Log.Message("InvoiceCurrency :"+InvoiceCurrency);
  
  
  //Changing Invoice Currency as Job Currency
    if(JobCurrency!=InvoiceCurrency){ 
    var prices =  Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 12)
    Sys.HighlightObject(prices);
    prices.Click(); 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
  
  aqUtils.Delay(2000, "Changing Job and Invoice Currency");
  var InvoiceCurr = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 11).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
  WorkspaceUtils.waitForObj(InvoiceCurr);
  InvoiceCurr.Keys(JobCurrency);
  aqUtils.Delay(4000, "Changing Invoice Currency as "+JobCurrency);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var Save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
  Sys.HighlightObject(Save)
  Save.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  }
  
  // Moving to Information Tab to Submit
  var info = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 5)
  Sys.HighlightObject(info);
  info.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  info.Click();
  count=false;
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  

  var Submit = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8);
  Sys.HighlightObject(Submit);
  Submit.Click();
  aqUtils.Delay(4000, "Submitting Job for Approval");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var Sliding_Panel = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
  Sys.HighlightObject(Sliding_Panel);
  Sliding_Panel.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  ImageRepository.ImageSet.Maximize.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var Job_Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
  Job_Approve.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
  
  var Approval_Table = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(Approval_Table);
    var y=0;
    
    //Getting User Name
    Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
  for(var i=0;i<Approval_Table.getItemCount();i++){   
     var approvers="";
      if(Approval_Table.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
        
      //Self Approve is Disabled. So finding different Approver
        var mainApprover = Approval_Table.getItem(i).getText_2(7).OleValue.toString().trim();
        var substitur = Approval_Table.getItem(i).getText_2(8).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+JobID+"*"+ temp;
      Log.Message("Approver level :" +i+ ": " +approvers);
      Approve_Level[y] = approvers;
      y++;
      }
}
TextUtils.writeLog("Finding approvers for Created Job");
CredentialLogin();
}
}



//Finding UserName for Approvers in Datasheets
function CredentialLogin(){ 
  var AppvLevl = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("FP")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("FP")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }

}


//Refreshing To-Do's List and Seleting Notification of Jobs
function todo(lvl,JobID,Apvr){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  
  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  
  
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}


if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}


var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job from To-Dos List"); 
listPass = false; 

//Finding Job in Approve Job Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
}


if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
  Log.Message(temp);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job (Substitute)").OleValue.toString().trim()+" (")
  Log.Message(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job (Substitute)").OleValue.toString().trim()+" (")!=-1)
  Log.Message(temp1.length==3)
  Log.Message(" ");
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job (Substitute) from To-Dos List");  
var listPass = false;   
//Finding Job in Approve Job (Substitute) Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
}  



if((listPass)||(!flag)){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job by Type from To-Dos List"); 
listPass = false; 

//Finding Job in Approve Job by Type Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
}


if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
  
  Log.Message(temp);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type (Substitute)").OleValue.toString().trim()+" (")
  Log.Message(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)
  Log.Message(temp1.length==3)
  Log.Message(" ");
  
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job by Type (Substitute) from To-Dos List"); 
var listPass = false;   

//Finding Job in Approve Job by Type (Substitute) Notification
FinalApproveJob(JobID,Apvr,lvl)
  }
} 
  }
  
}

//Finding Job is available in Approvers
function FinalApproveJob(JobID,Apvr,lvl){ 
  

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
//Finding Screen with Close Filter or Show Filter
var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
var showFilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
WorkspaceUtils.waitForObj(showFilter);
showFilter.Click();

}


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(table);
var firstCell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
firstCell.setText(JobID);
var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(5000, Indicator.Text);;
var i=0;
var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);

flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==JobID){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}


}




//Approve the Created Job
function ApproveJob(JobID,Apvr,lvl){ 
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }


if(flag){
vID_Status = false;
ValidationUtils.verify(flag,true,"Created Job is available in Approval List");
TextUtils.writeLog("Created Job is available in Approval List");
var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;

closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var Approve = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SWTObject("SingleToolItemControl", "", 9);
Sys.HighlightObject(Approve)

Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Sys.HighlightObject(Approve)
Approve.Click();


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(2000, "Approving Job");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
ValidationUtils.verify(true,true,"Job is Approved by "+Apvr)
TextUtils.writeLog("Job is Approved by "+Apvr);

// After Final Approve validating in Sliding panel
if(Approve_Level.length==lvl+1){
var approvalBar = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();

  aqUtils.Delay(2000, "Approving Job");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  ImageRepository.ImageSet.Maximize.Click();
    

  aqUtils.Delay(2000, "Approving Job");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var JobApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2

WorkspaceUtils.waitForObj(JobApproval);
JobApproval.HoverMouse();
ReportUtils.logStep_Screenshot();
JobApproval.Click();
aqUtils.Delay(3000, Indicator.Text);;

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

WorkspaceUtils.waitForObj(approvertable);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
ReportUtils.logStep_Screenshot();

for(var i=0;i<approvertable.getItemCount();i++){   
var approvers="";
if(approvertable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(true,false,"Job is not Approved in Level : "+i);
}
}

var JobApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl;
Sys.HighlightObject(JobApproval);
JobApproval.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

ImageRepository.ImageSet.Forward.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var Blanket_Invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Estimating = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Time_Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Purchase_Ordering = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Amount_Registration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Invoicing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 9).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");

//var Blanket_Invoice = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Estimating = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Time_Registration = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Purchase_Ordering = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Amount_Registration = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//var Invoicing = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.SWTObject("Composite", "", 9).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");

Sys.HighlightObject(Blanket_Invoice)
Sys.HighlightObject(Estimating)
Sys.HighlightObject(Time_Registration)
Sys.HighlightObject(Purchase_Ordering)
Sys.HighlightObject(Amount_Registration)
Sys.HighlightObject(Invoicing)

var checkmark = false;

    if(Blanket_Invoice.getSelection()){
      Blanket_Invoice.Click();
      ValidationUtils.verify(true,true,"Job for Blanket Invoice is Unchecked");
      checkmark = true;
    } 
    
    if(Estimating.getSelection()){
      Estimating.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Estimating is Unchecked");
      checkmark = true;
    } 
    
    if(Time_Registration.getSelection()){
      Time_Registration.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Time Registration is Unchecked");
      checkmark = true;
    } 
    
    if(Purchase_Ordering.getSelection()){
      Purchase_Ordering.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Purchase Ordering is Unchecked");
      checkmark = true;
    } 
    
     if(Amount_Registration.getSelection()){
      Amount_Registration.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Amount Registration is Unchecked");
      checkmark = true;
    } 
    
     if(Invoicing.getSelection()){
      Invoicing.Click();
      ValidationUtils.verify(true,true,"Job Blocked for Invoicing is Unchecked");
      checkmark = true;
    } 
    
    if(checkmark){ 
      aqUtils.Delay(3000, "Saving Changes for Job");;
      var Save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      Sys.HighlightObject(Save);
      Save.Click();
      aqUtils.Delay(3000, "Saving Changes for Job");;
    }
    
//var Save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//      Sys.HighlightObject(Save);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}

  ValidationUtils.verify(true,true,"Job is Approved by "+Apvr)

}


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


function ConvetToQuote(){ 
  

    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(3000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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

  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);



  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  WorkspaceUtils.waitForObj(job);
  job.Click();

  job.setText(JobID);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(4000, "Reading Table Data in Job List");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  //Finding Created Job
  flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==JobID){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  
  
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  ReportUtils.logStep_Screenshot("");
  closeFilter.Click();
  aqUtils.Delay(4000, "Created Job Details is loading");
  
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  //Validating Job 
  aqUtils.Delay(2000, "Checking Job ");

  }
  // Moving to Information Tab to Convert To Quote
  var info = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 5)
  Sys.HighlightObject(info);
  info.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  info.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  //Validating Job 
  aqUtils.Delay(2000, "Checking Job ");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var Convert_To_Qupte = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SWTObject("SingleToolItemControl", "", 13);
  Sys.HighlightObject(Convert_To_Qupte);
  Convert_To_Qupte.Click();
  aqUtils.Delay(2000, "Checking Job ");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  }
}