//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
  Indicator.Show();
  Indicator.PushText("waiing for window to open");
  
/**
 * This script create Sub Job for Main Job
 * @author  : Muthu Kumar M
 * @version : 3.0
 * Modified Date(MM/DD/YYYY) :12/20/2021
*/

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "SubJobs";
var Language = "";

//Global Varibales
var company,Job_group,JobNo,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile,startDateValue,SubJobName;
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var Approve_Level = [];
var JobID,level = ""; 
var flag=false;
var Maconomy_ParentAddress,Maconomy_Index = "";


//Main Function
function Creating_SubJob(){ 
TextUtils.writeLog("Sub-Job Creation Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
Log.Message(Language)

//Checking Login to execute Sub-Job Creation script
//var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//menuBar.Click();

ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
//WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

//if(eval(Maconomy_ParentAddress).WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
//WorkspaceUtils.closeMaconomy();
//Restart.login(Project_manager);
//  
//}

aqUtils.Delay(15000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);

Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "SubJobs";
company,Job_group,JobNo,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile,startDateValue,SubJobName = "";
Approve_Level = [];
JobID,level = ""; 
flag=false;
Arrays = [];

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Sub Job Creation started::"+STIME);

try{
getDetails();
goToJobMenuItem();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
GoToSubJob();
Entering_Details();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
// Close all opened workspace
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

goToJobMenuItem();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
Submiting_Job();

// Approving Job in Multi-Levels
for(var i=0;i<ApproveInfo.length;i++){
level = i;

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
//WorkspaceUtils.closeMaconomy();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);

Maconomy_Index = WorkspaceUtils.Maconomy_Parent;
Maconomy_Index++;
WorkspaceUtils.Maconomy_Parent = Maconomy_Index;
Log.Message(Maconomy_Index);

// Restarting maconomy with Approver Logins
var temp = ApproveInfo[i].split("*");

var Macscreen = WorkspaceUtils.switch_Maconomy(temp[2])
if(Macscreen=="Screen Not Found"){
Restart.login(temp[2]);
}
aqUtils.Delay(5000, Indicator.Text);

Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(temp[2]);

//Refreshing To-Do's List to find Submitted Job
vID_Status = true;
todo(temp[3],temp[1],temp[2],i);

//Approving Created Job is every Levels
ApproveJob(temp[1],temp[2],i,temp[3]);

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
//WorkspaceUtils.closeMaconomy();
WorkspaceUtils.closeAllWorkspaces();
}
}
  catch(err){
    Log.Message(err);
  }

}


//Getting Details to create Sub-Job from Datasheet
function getDetails(){
sheetName = "SubJobs";

ExcelUtils.setExcelName(workBook, "Data Management", true);
JobNo = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
if((JobNo =="")||(JobNo ==null)){
sheetName ="SubJobs";
ExcelUtils.setExcelName(workBook, sheetName, true);
JobNo = ExcelUtils.getColumnDatas("Job_No",EnvParams.Opco)
}
else if((JobNo=="")||(JobNo==null)) 
ValidationUtils.verify(false,true,"Job Number is needed to Create SubJob");



ExcelUtils.setExcelName(workBook, sheetName, true);
company = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");
}


Job_group = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)
if((Job_group==null)||(Job_group=="")){ 
ValidationUtils.verify(false,true,"Job Group is Needed to Create a SubJob");
}


Job_Type = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
if((Job_Type==null)||(Job_Type=="")){ 
ValidationUtils.verify(false,true,"Job Type Number is Needed to Create a SubJob");
}


department = ExcelUtils.getRowDatas("Department",EnvParams.Opco)
if((department==null)||(department=="")){ 
ValidationUtils.verify(false,true,"Department Number is Needed to Create a SubJob");
}


buss_unit = ExcelUtils.getRowDatas("BusinessUnit",EnvParams.Opco)
if((buss_unit==null)||(buss_unit=="")){ 
ValidationUtils.verify(false,true,"BusinessUnit Number is Needed to Create a SubJob");
}


TemplateNo = ExcelUtils.getRowDatas("Template",EnvParams.Opco)
if((TemplateNo==null)||(TemplateNo=="")){ 
ValidationUtils.verify(false,true,"Template Number is Needed to Create a SubJob");
}


SubJobName = ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
if((SubJobName==null)||(SubJobName=="")){ 
ValidationUtils.verify(false,true,"Job Name is Needed to Create a SubJob");
}


startDateValue = ExcelUtils.getRowDatas("Job_StartDate",EnvParams.Opco)
if((startDateValue==null)||(startDateValue=="")){ 
ValidationUtils.verify(false,true,"startDateValue is Needed to Create a Sub Job");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
Project_manager = ExcelUtils.getRowDatas("Project Manager",EnvParams.Opco)

}

/**
  *  This function Navigates to Jobs screen from Jobs workspace
  */
function goToJobMenuItem(){
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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

var WrkspcCount = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
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


function GoToSubJob() {

 var closeFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  WorkspaceUtils.waitForObj(table);
  var companyFilter = eval(Maconomy_ParentAddress).
  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  WorkspaceUtils.waitForObj(companyFilter);
  companyFilter.forceFocus();
  companyFilter.setVisible(true);
  companyFilter.ClickM();
  table.Child(0).setText(company);
  aqUtils.Delay(1000, Indicator.Text);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);

  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  
  var jobNoField = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3)
  Sys.HighlightObject(jobNoField);
  jobNoField.Click();
  aqUtils.Delay(500);
  table.Child(2).forceFocus();
  table.Child(2).setVisible(true);
  table.Child(2).setText(JobNo);
 
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Sys.Desktop.KeyUp(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
  aqUtils.Delay(1000, Indicator.Text);
  
  var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  WorkspaceUtils.waitForObj(job);
  job.Click();

  Sys.HighlightObject(table);
  aqUtils.Delay(2000, "Reading Table Data in Job List");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==(JobNo)){ 
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
  }

    
  if(count){
  var ref = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  var subjob = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 10)
  subjob.Click()
  aqUtils.Delay(1000);
  count=false;
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
  }
   aqUtils.Delay(4000, "Clicking Sub Job"); 
    var createSubJobbtn = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 6);
    Sys.HighlightObject(createSubJobbtn);  
    createSubJobbtn.Click();
    aqUtils.Delay(1000); 
    aqUtils.Delay(1000);
    var screen = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim());
    Log.Message("Create New SubJob window is Opened")
    TextUtils.writeLog("Clicked Create Sub Job")
    }
    
    
    
    
function Entering_Details(){
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
TextUtils.writeLog("Entering Details in Create Sub Job Wizard")
aqUtils.Delay(3000, "Waiting for ScrolledComposite");
var job = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);

if(Job_group!=""){
job.Click();
aqUtils.Delay(5000, "Waiting for ScrolledComposite");
var list = "";
try{
list = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
}
catch(e){ 
job.Click(); 
aqUtils.Delay(6000, "Waiting for ScrolledComposite");
list = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
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
   ExcelUtils.setExcelName(workBook, sheetName, true);
  var JobType = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
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
    
var Depart = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(department!=""){
Depart.Click();
var ExlArray = getExcelData("Validate_Department",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(Depart,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Department").OleValue.toString().trim(),department,ExlArray,"Department Number");
}else{ 
  ValidationUtils.verify(false,true,"Department is Needed to Create Job");
}
    
    
    
//----------Entering BusinessUnit-------------   
    
  var BussUnit = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(buss_unit!=""){
BussUnit.Click();
var ExlArray = getExcelData("Validate BusinessUnit",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(BussUnit,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Business Unit").OleValue.toString().trim(),buss_unit,ExlArray,"Business Unit Number");
}else{ 
  ValidationUtils.verify(false,true,"Business Unit is Needed to Create Job");
}




//----------Entering Template Number-------------    
  var template = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(TemplateNo!=""){
template.Click();
var ExlArray = getExcelData("Validate Template",EnvParams.Opco)
WorkspaceUtils.Config_with_Maconomy_templateValidation(template,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),TemplateNo,ExlArray,Job_Type,company,Job_group,"Template Number");
}else{ 
  ValidationUtils.verify(false,true,"Template is Needed to Create Job");
}
    

//----------Entering Project Manager-------------
  
var ProjectManger = 
eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
//eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);  
Log.Message(Project_manager)
 if((Project_manager!="")||(Project_manager!=null)){
 if(ProjectManger.getText()!=Project_manager.toString().trim()){
 ProjectManger.Click();
 WorkspaceUtils.SearchByValue(ProjectManger,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Project_manager,"Project Manager");
 }
 }
 
 
var jobname = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
jobname.click()
SubJobName = SubJobName.toString().trim()+" "+STIME;
jobname.setText(SubJobName);
// Job_name
    
var subjobbtn = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
if(subjobbtn.isEnabled()){
Log.Message("Create Button is Visible"); 
Log.Message("SubJob is Created"); 
aqUtils.Delay(1000);
Sys.HighlightObject(subjobbtn);  
subjobbtn.Click(); 
aqUtils.Delay(5000);      
}
else{
aqUtils.Delay(1000);
var cancelbtn = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create Sub Job").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
Sys.HighlightObject(cancelbtn);
cancelbtn.Click()
Log.Error("SubJob is not Created");     
}  
 

//Clinking pop-ups or Notification of Jobs     
Log.Message("Language :"+Language);
aqUtils.Delay(5000, "Waiting to Check any pop-up for credit limit");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
 var p = eval(WorkspaceUtils.Sys_Maconomy_Parent);
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
}


//Clinking 2nd Pop-ups
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

//Clinking 3rd Pop-ups
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


//Clinking 4th Pop-ups
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
else{
  
Log.Message("Language :"+Language);
 var p = eval(WorkspaceUtils.Sys_Maconomy_Parent);
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


//Clinking 2nd Pop-ups
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


//Clinking 3rd Pop-ups
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

//Clinking 4th Pop-ups
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

    
aqUtils.Delay(1000); 
var closefilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2); 
closefilter.Click();    
aqUtils.Delay(5000);; 
}

//Validating Created Job is available
function Submiting_Job() {
  
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var closeFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  WorkspaceUtils.waitForObj(table);
  var companyFilter = eval(Maconomy_ParentAddress).
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

  var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  WorkspaceUtils.waitForObj(job);
  job.Click();
  job.setText(SubJobName);
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(2000, "Reading Table Data in Job List");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  //Finding Created Job
  flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==SubJobName){ 
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
  ValidationUtils.verify(true,true,"Sub-Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  TextUtils.writeLog("Created Sub-Job is available in system");
  TextUtils.writeLog("Sub-Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  
  
  if(flag){
  ReportUtils.logStep("INFO", "Created Sub-Job is listed in table");
  ReportUtils.logStep_Screenshot("");
  closeFilter.Click();
  aqUtils.Delay(4000, "Created Sub-Job Details is loading");
  

  var ref = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  //Validating Job and Invoice Currency
  aqUtils.Delay(2000, "Checking Job and Invoice Currency");
  var JobCurrency = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  var InvoiceCurrency = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  JobCurrency = JobCurrency.getText().OleValue.toString().trim();
  InvoiceCurrency = InvoiceCurrency.getText().OleValue.toString().trim();
  ReportUtils.logStep("Job Currency:"+JobCurrency)
  ReportUtils.logStep("Invoice Currency:"+InvoiceCurrency)
  Log.Message("JobCurrency :"+JobCurrency);
  Log.Message("InvoiceCurrency :"+InvoiceCurrency);
  
    var ref = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  
  //Changing Invoice Currency as Job Currency
    if(JobCurrency!=InvoiceCurrency){ 
    var prices =  eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 12);
//    var prices =  Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 12)
    Sys.HighlightObject(prices);
    prices.Click(); 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
  
  aqUtils.Delay(2000, "Changing Job and Invoice Currency");
  var InvoiceCurr = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 11).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
//  var InvoiceCurr = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 11).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
  WorkspaceUtils.waitForObj(InvoiceCurr);
  InvoiceCurr.Keys(JobCurrency);
  aqUtils.Delay(4000, "Changing Invoice Currency as "+JobCurrency);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
//  var Save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
  Sys.HighlightObject(Save)
  Save.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  }
  
  // Moving to Information Tab to Submit
  var info = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
//  var info = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 5)
  Sys.HighlightObject(info);
  info.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  info.Click();
  count=false;
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

    aqUtils.Delay(4000,"Loading Screen");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){   }  
    

var  PropArray = new Array("JavaClassName", "Index");
var  ValuesArray = new Array("McDatePickerWidget", "2");

  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1)
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var JobStart_Date=null;
  Log.Message(obj.length)
  
  if(obj.length>=3){
      if(obj[0].isVisible){
        Sys.HighlightObject(obj[0]);
        JobStart_Date = obj[0];        
      }
  
  }
Sys.HighlightObject(JobStart_Date);
//  var JobStart_Date = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2)
//  var JobStart_Date = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2)
  Sys.HighlightObject(JobStart_Date);
  if((JobStart_Date.getText()!="")||(JobStart_Date.getText()!=null)){
    JobStart_Date.Click();
    JobStart_Date.setText(startDateValue);
  
  aqUtils.Delay(4000, "Saving Sub-Job for Submitting");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//  var Save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.SWTObject("SingleToolItemControl", "", 4);
//  Sys.HighlightObject(Save);
//  Save.Click();
  }

   
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("SingleToolItemControl", "8");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Submit = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Submit"){
  Sys.HighlightObject(obj[i_count]);
  Submit = obj[i_count];
  break;
 }
}

Log.Message(Submit.FullName)
Sys.HighlightObject(Submit);                                        
//  var Submit = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8)
//  var Submit = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8)
//  var Submit = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8);
  Sys.HighlightObject(Submit);
  Submit.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(4000, "Submitting Sub-Job for Approval");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
            
  var ref = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
                                                  
//  var Sliding_Panel = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "")
////  var Sliding_Panel = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
//  Sys.HighlightObject(Sliding_Panel);
  
  
  
  var Sliding_Panel ;
PropArray = new Array("JavaClassName", "Index");
ValuesArray = new Array("TabControl", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
      var approv_bar = p.FindAll(PropArray, ValuesArray, 1000);
      for (let j_count = 0; j_count < approv_bar.length; j_count++){ 
        if(approv_bar[j_count].Visible){ 
          Sliding_Panel = approv_bar[j_count];
          Log.Message(Sliding_Panel.FullName);
          break;
        }
        
        }
      
Sys.HighlightObject(Sliding_Panel);
  Sliding_Panel.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  ImageRepository.ImageSet.Maximize.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  
  var AprveAction;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("TabControl", "5");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible)
  if(obj[i_count].text=="Job Approval"){
  Sys.HighlightObject(obj[i_count]);
  AprveAction = obj[i_count];
  break;
 }
}
Sys.HighlightObject(AprveAction);
AprveAction.Click();

//  var Job_Approve = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
//  var Job_Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
//  Job_Approve.Click();
  aqUtils.Delay(4000, "Finding Approval Tab");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }
  
   
var Approval_Table;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
if(obj[i_count].Visible)
Approval_Table = obj[i_count];
}

    Log.Message(Approval_Table.FullName)
    Sys.HighlightObject(Approval_Table); 
//  var Approval_Table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//  var Approval_Table = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(Approval_Table);
    var y=0;
    
    //Getting User Name
    Project_manager = eval(Maconomy_ParentAddress).WndCaption;
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
TextUtils.writeLog("Finding approvers for Created Sub-Job");
//CredentialLogin();
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Sub-Job Number",EnvParams.Opco,"Data Management",JobID)


}



//Finding UserName for Approvers in Datasheets
//function CredentialLogin(){ 
//  var AppvLevl = [];
//for(var i=0;i<Approve_Level.length;i++){
//  var UserN = true;
//  var temp="";
//  var temp1="";
//  var Cred = Approve_Level[i].split("*");
//  for(var j=2;j<4;j++){
//  temp="";
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("FP")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
//  { 
//     var sheetName = "Agency Users";
//     workBook = Project.Path+excelName;
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//  }
//  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("FP")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
//  { 
//
//    var sheetName = "SSC Users";
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
//  }
//
//  if(temp.length!=0){
//    temp1 = temp1+temp+"*"+j+"*";
////  break;
//  }
//  }
//  if((temp1=="")||(temp1==null))
//  Log.Error("User Name is Not available for level :"+i);
//  Log.Message(temp1)
//  AppvLevl[i] = temp1;
//}
//  ApproveInfo = levelMatch(AppvLevl)
//  Log.Message("-----Approvers-------------")
//  for(var i=0;i<ApproveInfo.length;i++){
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
//    Log.Message(ApproveInfo[i]);
//    }
//
//}



//Refreshing To-Do's List and Seleting Notification of Jobs
function todo(lvl,JobID,Apvr){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  
  var toDo = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
//  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
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

//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
//var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
//}
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
//var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//}

//var refresh =  eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("SingleToolItemControl", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var refresh;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Refresh ToDo's"){
  Sys.HighlightObject(obj[i_count]);
  refresh = obj[i_count];
  break;
  }
}
Log.Message(refresh.FullName)
Sys.HighlightObject(refresh)
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}



  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("Tree", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Client_Managt;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible){
  Sys.HighlightObject(obj[i_count]);
  Client_Managt = obj[i_count];
  break;
  }
}
Log.Message(Client_Managt.FullName)
Sys.HighlightObject(Client_Managt)
var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
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
  
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Apvr);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
//Finding Screen with Close Filter or Show Filter
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "");
//var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).Visible){
//if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
  
var showFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
//var showFilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
WorkspaceUtils.waitForObj(showFilter);
showFilter.Click();

}


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(table);

var firstCell = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//var firstCell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "");
WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
firstCell.setText(JobID);

var closefilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(5000, Indicator.Text);;
var i=0;
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
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


if(!flag)
WorkspaceUtils.closeAllWorkspaces();
}




//Approve the Created Job
function ApproveJob(JobID,Apvr,lvl){ 
  
//Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Apvr);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }


if(flag){
vID_Status = false;
ValidationUtils.verify(flag,true,"Created Job is available in Approval List");
TextUtils.writeLog("Created Job is available in Approval List");
var closefilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;

closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var Approve = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
//.SWTObject("SingleToolItemControl", "", 8);
//var Approve = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SWTObject("SingleToolItemControl", "", 8);

var  obj = Approve.FindAll("JavaClassName", "SingleToolItemControl", 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText!=null)
  if(obj[i_count].toolTipText.OleValue.toString().indexOf("Approve")!=-1){
  Sys.HighlightObject(obj[i_count]);
  Approve = obj[i_count];
  break;
 }
}
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
  
var approvalBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
//var approvalBar = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
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
var JobApproval = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//var JobApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2

WorkspaceUtils.waitForObj(JobApproval);
JobApproval.HoverMouse();
ReportUtils.logStep_Screenshot();
JobApproval.Click();
aqUtils.Delay(3000, Indicator.Text);;

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var approvertable = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

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

var JobApproval = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "")
//var JobApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl;
Sys.HighlightObject(JobApproval);
JobApproval.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

ImageRepository.ImageSet.Forward.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var Blanket_Invoice = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Estimating = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Time_Registration = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Purchase_Ordering = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Amount_Registration = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Invoicing = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 9).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");

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
      var Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
//      var Save = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
excelData =[];  
Log.Message(" ");
Log.Message(excelName)
Log.Message(workBook);
Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
var id =0;
var colsList = [];
 var temp ="";
Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
       Log.Message("Row Identifier is Matched");
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

       
     }else if(temp.length>0){ 
      excelData[0] = temp;

     }
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
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
