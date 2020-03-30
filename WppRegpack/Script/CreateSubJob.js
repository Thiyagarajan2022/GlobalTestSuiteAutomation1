//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT EnvParams
  Indicator.Show();
  Indicator.PushText("waiing for window to open");

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "SubJobs";
Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";


var company,Job_group,JobNo,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile,startDateValue;


function getDetails(){
  
ExcelUtils.setExcelName(workBook, sheetName, true);
company = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");
}

Log.Message(company);

JobNo = ExcelUtils.getRowDatas("Job_No",EnvParams.Opco)
if((JobNo==null)||(JobNo=="")){ 
JobNo =readlog();
Log.Message("jobNoNotepad= "+JobNo);
}
if((JobNo==null)||(JobNo=="")){ 
ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");
}

Log.Message(JobNo);

Job_group = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)
if((Job_group==null)||(Job_group=="")){ 
ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");
}
Log.Message(Job_group);

Job_Type = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
if((Job_Type==null)||(Job_Type=="")){ 
ValidationUtils.verify(false,true,"Job Type Number is Needed to Create a Job");
}
Log.Message(Job_Type);

department = ExcelUtils.getRowDatas("Department",EnvParams.Opco)
if((department==null)||(department=="")){ 
ValidationUtils.verify(false,true,"Department Number is Needed to Create a Job");
}
Log.Message(department);

buss_unit = ExcelUtils.getRowDatas("BusinessUnit",EnvParams.Opco)
if((buss_unit==null)||(buss_unit=="")){ 
ValidationUtils.verify(false,true,"BusinessUnit Number is Needed to Create a Job");
}
Log.Message(buss_unit);

TemplateNo = ExcelUtils.getRowDatas("Template",EnvParams.Opco)
if((TemplateNo==null)||(TemplateNo=="")){ 
ValidationUtils.verify(false,true,"Template Number is Needed to Create a Job");
}
Log.Message(TemplateNo);

Job_name = ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
if((Job_name==null)||(Job_name=="")){ 
ValidationUtils.verify(false,true,"Job Name is Needed to Create a Job");
}
Log.Message(Job_name);

startDateValue = ExcelUtils.getRowDatas("Job_StartDate",EnvParams.Opco)
if((startDateValue==null)||(startDateValue=="")){ 
ValidationUtils.verify(false,true,"startDateValue is Needed to Create a Sub Job");
}
ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

}

function SubJobAddress(){ 
//Checking Labels in Job Create Wizard
Sys.Process("Maconomy").Refresh();

var job = 
Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
WorkspaceUtils.waitForObj(job);
if(job!="Job Group")
ValidationUtils.verify(false,true,"Job Group field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Job Group field is available in Maconomy for Job Creation");
var JobType = 
Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(JobType!="Job Type")
ValidationUtils.verify(false,true,"Job Type field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Job Type field is available in Maconomy for Job Creation");
var Depart = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Depart!="Department")
ValidationUtils.verify(false,true,"Department field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Department field is available in Maconomy for Job Creation");
var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(BussUnit!="Business Unit")
ValidationUtils.verify(false,true,"Business Unit field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Business Unit field is available in Maconomy for Job Creation");
var template = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(template!="Template")
ValidationUtils.verify(false,true,"Template field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Template field is available in Maconomy for Job Creation");
var prdNumber = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(prdNumber!="Product")
ValidationUtils.verify(false,true,"Product field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Product field is available in Maconomy for Job Creation");
var jobName = 
Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(jobName!="Job Name")
ValidationUtils.verify(false,true,"Job Name field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Job Name field is available in Maconomy for Job Creation");
var ProjectManger = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(ProjectManger!="Project Manager")
ValidationUtils.verify(false,true,"Project Manager field is missing in Maconomy for Job Creation");
else
ValidationUtils.verify(true,true,"Project Manager field is available in Maconomy for Job Creation");

}

function GoToCreatedSubjob() {
  
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WaitSWTObject("Composite", "",1,60000).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
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
    companyFilter.Click();
    table.Child(0).setText(company);
    aqUtils.Delay(500);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09); // Press Ctrl
    aqUtils.Delay(500);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(500);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);    
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyUp(0x09);
    var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3); 
    WorkspaceUtils.waitForObj(job);
    job.Click();
    aqUtils.Delay(500);
    table.Child(2).forceFocus();
    table.Child(2).setVisible(true);
    table.Child(2).setText(Job_name+" "+STIME);
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    aqUtils.Delay(500);; 

}

  function GoToSubJob() {

    aqUtils.Delay(500);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WaitSWTObject("Composite", "",1,60000).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    aqUtils.Delay(500);
    WorkspaceUtils.waitForObj(table);
    var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
    SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
    SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
    SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
    SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
    SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    companyFilter.forceFocus();
    WorkspaceUtils.waitForObj(companyFilter);
    companyFilter.setVisible(true);
    companyFilter.Click();
 //   table.Child(0).setText("^a[BS]");
    table.Child(0).setText(company);
    aqUtils.Delay(500);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09); // Press Ctrl
    aqUtils.Delay(500);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);    
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyUp(0x09);
    var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3)
    WorkspaceUtils.waitForObj(job);
    job.Click();
    aqUtils.Delay(500);
    table.Child(2).forceFocus();
    table.Child(2).setVisible(true);
    table.Child(2).setText(JobNo);
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    aqUtils.Delay(1000);
    
    
    if(count){
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    WorkspaceUtils.waitForObj(ref);
     ref.Refresh();
    var subjob = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 10)
   WorkspaceUtils.waitForObj(subjob);
     subjob.Click()
    aqUtils.Delay(1000);
    count=false;
    }
    
    var createSubJobbtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 6);
    WorkspaceUtils.waitForObj(createSubJobbtn);
    Sys.HighlightObject(createSubJobbtn);  
    createSubJobbtn.Click();
    aqUtils.Delay(1000); 
    aqUtils.Delay(1000);
    var screen = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job");
    Log.Message("Create New SubJob window is Opened")
    
    SubJobAddress();
    
//----------Entering Job Group-------------

var job = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);

if(Job_group!=""){
job.Click();
aqUtils.Delay(1000);;
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible2 = true;
while(Add_Visible2){
if(list.isEnabled()){
Add_Visible2 = false;
var dropstatus = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Job_group.toString().trim()){ 
          list.Keys("[Enter]");
          aqUtils.Delay(1000);
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
  var JobType = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Job_Type!=""){
JobType.Click();
  var JG = "";
  if(Job_group.toString().trim()=="Client Billable")
  JG = "ValidateJobtype_CB(Client Billable)";
  if(Job_group.toString().trim()=="Client Non-Billable")
  JG = "ValidateJobtype_CNB(Client Non-Billable)";
  if(Job_group.toString().trim()=="Internal")
  JG = "ValidateJobtype_IN(Internal)";
var ExlArray = []; 
Log.Message(EnvParams.Opco);
ExlArray = getExcelData(JG,EnvParams.Opco);
if(ExlArray.length>0){ 
  
}
else
ValidationUtils.verify(false,true,"Selected Job Group doesn't have any Job Type in Opco's");
Job_Type = WorkspaceUtils.config_with_Maconomy_Validation(JobType,"Job Type",Job_Type,ExlArray,"Job Type");
}else{ 
  ValidationUtils.verify(false,true,"JobType is Needed to Create Job");
}
     
//----------Entering Department-------------   
    
var Depart = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(department!=""){
Depart.Click();
var ExlArray = getExcelData("Validate_Department",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(Depart,"Department",department,ExlArray,"Department Number");
}else{ 
  ValidationUtils.verify(false,true,"Department is Needed to Create Job");
}
    
    
    
    //----------Entering BusinessUnit-------------   
    
  var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(buss_unit!=""){
BussUnit.Click();
var ExlArray = getExcelData("Validate BusinessUnit",EnvParams.Opco)
WorkspaceUtils.config_with_Maconomy_Validation(BussUnit,"Business Unit",buss_unit,ExlArray,"Business Unit Number");
}else{ 
  ValidationUtils.verify(false,true,"Business Unit is Needed to Create Job");
}

//----------Entering Template Number-------------    
  var template = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(TemplateNo!=""){
template.Click();
var ExlArray = getExcelData("Validate Template",EnvParams.Opco)
WorkspaceUtils.Config_with_Maconomy_templateValidation(template,"Job",TemplateNo,ExlArray,Job_Type,company,Job_group,"Template Number");
}else{ 
  ValidationUtils.verify(false,true,"Template is Needed to Create Job");
}
    

//----------Entering Project Manager-------------
  
var ProjectManger = 
Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
//Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);  
 Log.Message(Project_manager);
 Log.Message(ProjectManger.getText())
 if(Project_manager!=""){
 if(ProjectManger.getText()!=Project_manager.toString().trim()){
 ProjectManger.Click();
 WorkspaceUtils.SearchByValues_Wiz2_Col_2(ProjectManger,"Employee",Project_manager,"Project Manager");
 }
 }

   var jobname = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
    jobname.click()
    jobname.setText(Job_name.toString().trim()+" "+STIME);
    
    
    var subjobbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
    if(subjobbtn.isEnabled()){
      Log.Message("Create Button is Visible"); 
      Log.Message("SubJob is Created"); 
      aqUtils.Delay(1000);
      Sys.HighlightObject(subjobbtn);  
      subjobbtn.Click();     
      } 
    else{
      aqUtils.Delay(1000);
      var cancelbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
      Sys.HighlightObject(cancelbtn);
      cancelbtn.Click()
      Log.Error("SubJob is not Created");     
    } 
    aqUtils.Delay(1000); 
  
    var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2); 
    closefilter.Click();    
}

    function gotoinformation(){   
    
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    table.Click();
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyUp(0x09);
    table.setText(Job_name+" "+STIME);
    aqUtils.Delay(1000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    var flag=false;
    for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(Job_name+" "+STIME)){ 
        flag=true;
        break;
      }
      else{ 
        table.Keys("[Down]");
      }
     }
ValidationUtils.verify(flag,true,"Sub Job is Created and available in Maconomy"); 


  if(flag){
    var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  closeFilter.Click();
  aqUtils.Delay(3000, Indicator.Text);

count = true;
  aqUtils.Delay(1000);

  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
 var info = 
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
info.Click();
  count=false;
  }
  aqUtils.Delay(1000);
  var Templete_Job = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
 Sys.HighlightObject(Templete_Job);

 var Blanket_invoice = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  Sys.HighlightObject(Blanket_invoice);

 var estimation = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
//  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
 Sys.HighlightObject(estimation);

  var Amount_Registrations = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  Sys.HighlightObject(Amount_Registrations);

 var Invoicing = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
 // Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", ""); 
  Sys.HighlightObject(Invoicing);

  var TimeReg = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  Sys.HighlightObject(TimeReg);

   var invoicing =Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPlainCheckboxView.InvoiceingCheckbox;
  Sys.HighlightObject(invoicing);
  checkmark = false;
  
  var jobtypecode = Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.Composite.jobtypecode
  Sys.HighlightObject(jobtypecode);
  var departmenttypecode = Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite2.Composite.jobDepartmentCode;
  Sys.HighlightObject(departmenttypecode);
  var businessunitcode =Aliases.Maconomy.SubJob.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite3.Composite.BusinessUnitCode
  //Aliases.ObjectGroup.businessUnitCode;
  Sys.HighlightObject(businessunitcode);
    
    
        if(jobtypecode.getText().trim()  !="")
        {
        ReportUtils.logStep("INFO", "Jobtype code is present");
        ValidationUtils.verify(true,true,"Jobtype code is present");
        Log.Checkpoint("jobtype code is present")
        }
        else{
        ValidationUtils.verify(false,true,"Jobtype code is present");
        }
        
        if(departmenttypecode.getText().trim()!="")
        {
        ReportUtils.logStep("INFO", "Departmenttype code is present");
        ValidationUtils.verify(true,true,"Departmenttype code is present");
        Log.Checkpoint("Departmenttype code is present")
        }
        else{
           ValidationUtils.verify(false,true,"Departmenttypecode is present");
        }
        
        
        if(businessunitcode.getText().trim()!="")
        {
        ReportUtils.logStep("INFO", "Businessunit code is present");
        ValidationUtils.verify(true,true,"Businessunit code is present");
        Log.Checkpoint("Businessunit code is present")
        }
        else{
        ValidationUtils.verify(false,true,"Businessunit code is present");
        }
  
  
    if(Templete_Job.getSelection()){ 
    Trmplete_Job.Click();
      Log.Message("Trmplete_Job is UnChecked")
      checkmark = true;
    }
    if(Blanket_invoice.getSelection()){ 
    Blanket_invoice.Click();
      Log.Message("Blanket is UnChecked")
      checkmark = true;
    }
    if(estimation.getSelection()){
      estimation.Click();
      Log.Message("estimating is Unchecked")
      checkmark = true;
    }     
    if(TimeReg.getSelection()){ 
    TimeReg.Click();
      Log.Message("TimeReg is UnChecked")
      checkmark = true;
    }
    if(Amount_Registrations.getSelection()){ 
    Amount_Registrations.Click();
      Log.Message("Amount_Registrations is UnChecked")
      checkmark = true;
    }
    if(Invoicing.getSelection()){
      Invoicing.Click();
      Log.Message("Invocing is Unchecked")
      checkmark = true;
    }
    
     if(invoicing.getSelection()){
      invoicing.Click();
      Log.Message("Invocing is Unchecked")
      checkmark = true;
    }
    
    var startdate = 
    
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2);
  startdate.setText(startDateValue);

  
  if(checkmark){ 
      
    aqUtils.Delay(1000);
    var savebtn =
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4).toolTipText;
    Log.Message("savebtn"+savebtn);
    
      var ref1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref1.Refresh();
  
        var savebtn1 =   
        
         Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(savebtn1);
    savebtn1.Click();
    Log.Checkpoint("Changes is Saved");
   aqUtils.Delay(1000);    
   if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Jobs - Information"){
   var OK = Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
   aqUtils.Delay(1000);
   OK.Click();
   }
    
  }
  
}
  }   
  
   
 function test()
 {
   
   var savebtn1 =   
        
         Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(savebtn1);
 }



function goToJobMenuItem(){

    var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
    menuBar.DblClick();
      if(ImageRepository.ImageSet.Jobs.Exists()){
          ImageRepository.ImageSet.Jobs.Click();// GL
      }
      else if(ImageRepository.ImageSet.Job.Exists()){
          ImageRepository.ImageSet.Job.Click();
      }
      else{
          ImageRepository.ImageSet.Jobs1.Click();
      }
      
    var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
      var Client_Managt;
    for(var i=1;i<=childCC;i++){ 
    Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
    if(Client_Managt.isVisible()){ 
    Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
    Client_Managt.DblClickItem("|Jobs");
    }
    }
       aqUtils.Delay(1000);

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

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}

function readlog(){

sheetName = "SubJobs";

ExcelUtils.setExcelName(workBook, sheetName, true);

comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)

if((comapany==null)||(comapany=="")){

ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");

}

Job_group = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)

if((Job_group==null)||(Job_group=="")){

ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");

}

var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)

var name =  LogReport_name(ExlArray,comapany,Job_group);

var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";

//Log.Message("Notepad :"+notepadPath)

return TextUtils.readDetails(notepadPath,"Job Number");

//Log.Message( readDetails("C:\\Users\\674087\\Documents\\TestComplete 14 Projects\\After Stuart Discussion\\WppRegression_v12.50\\WppRegPack\\RegressionLogs\\TESTAPAC\\Regression\\China\\1307-Sudler China(MDS)_Client Billable.txt","Job Number") );

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

//Log.Message(compStatus);

return compStatus

}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
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

function SubJob(){ 
Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Log.Message(EnvParams.Opco)
Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Log.Message(Language)
if(Language=="English"){
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Sub Job Creation started::"+STIME);

    getDetails();
    goToJobMenuItem();
    GoToSubJob();
    gotoinformation();
    closeAllWorkspaces();
}
else{ 
 JobCreation.SpanishcreateJob();
}
}



