//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT ObjectUtils
//USEUNIT ActionUtils

/**
 * This script create Quote and Client Approved Estimate for Main Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :02/10/2021
 * Modified Date(MM/DD/YYYY) : 01/19/2022
*/


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
 

//Global Varibales
var sheetName ="";
var STIME = "";
var comapany,Job_group = "";
var jobNumber = "";
var Approve_Level = [];
var y=0;
var ApproveInfo = [];
var level =0;
var Language = "";
var Estimate = [];
var ExcelEstimate = [];
var workEstimate = [];
var clientEstimate = [];
var RevisionNo = "";
var Language = "";
var C_Currency = "";
var QuoteDetails = [];
var Jobnum = "";
var Project_manager = ""
var Maconomy_ParentAddress,Maconomy_Index = "";

//Main Function
function CreateQuote(sheet,Jobnumber){ 
TextUtils.writeLog("Job Quote and Client Approved Estimate Creation Started"); 
Indicator.PushText("waiting for window to open");


//Getting Language from EnvParamaters.xlsx
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;


//Checking Login to execute Create Job Quotation
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);


//Re-Intialize Variable
excelName = EnvParams.path;
workBook = Project.Path+excelName;
STIME = "";
comapany,Job_group = "";
jobNumber,RevisionNo = "";
Approve_Level = [];
y=0;
ApproveInfo = [];
level =0;

comapany = EnvParams.Opco;
sheetName =sheet;
Jobnum = Jobnumber;


//try{
getDetails();
goTo_Job_Menu();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
goToBudget();
validatingWorkEstimate();
transferToQuote();

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);

ApproveQuote();

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

 Approve_Level = [];
 y=0;
 ApproveInfo = [];
 level =0;
 
 
//Checking Login to execute Create Job Quotation
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);
goTo_Job_Menu();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
SelectJob();
convertToOrder();
clientApprovedEsimate();
validatingclientEstimate();
WorkspaceUtils.closeAllWorkspaces();

// Approving Created Job Quote
for(var i=level;i<ApproveInfo.length;i++){
  level = i;
  
  

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
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
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
aprvBudget(temp[0],temp[1],temp[2]);

}
//}
//  catch(err){
//    Log.Message(err);
//  }

//Close all Open Workspace in Maconomy
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){ 
//sheetName ="CreateQuote";  

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number_"+Jobnum,EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Quote");
  

  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  RevisionNo = ReadExcelSheet("Budget Revision No",EnvParams.Opco,"Data Management");
  if((RevisionNo=="")||(RevisionNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  RevisionNo = ExcelUtils.getColumnDatas("Revision",EnvParams.Opco)
  }
  if((RevisionNo=="")||(RevisionNo==null))
  ValidationUtils.verify(false,true,"Revision Number is needed to Create Quote");
}



function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}  
    

/**
  *  This function Navigates to Jobs screen from Jobs workspace
  */
  
// Navigating to Jobs from Jobs Menu
function goTo_Job_Menu(){

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_Jobs_from_workspace(); //Select Jobs Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());

}



//Validating Working Estimate is Approved or Not
function goToBudget(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  
  var allJobs = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  
  var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  
  var firstcell = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  
  var closeFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
  
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
    
  ReportUtils.logStep("INFO", "Job is listed in table to create Quote");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job is available in maconommy to create Quote"); 
  closeFilter.Click();
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("TabControl", "5");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Budget;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Budgeting"){
  Sys.HighlightObject(obj[i_count]);
  Budget = obj[i_count];
  break;
  }
}

Log.Message(Budget.FullName)
WorkspaceUtils.waitForObj(Budget);
Budget.Click();
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
  aqUtils.Delay(5000, Indicator.Text);
  
   var show_budget;   
      PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGroupWidget", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].ChildCount>=8))
show_budget = obj[i_count];
}
    Sys.HighlightObject(show_budget);
show_budget = show_budget.SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);
    WorkspaceUtils.waitForObj(show_budget);
    show_budget.HoverMouse();
    show_budget.HoverMouse();
    show_budget.HoverMouse();
    
    
Sys.HighlightObject(show_budget);    
WorkspaceUtils.waitForObj(show_budget);
show_budget.Keys("Working Estimate"); 
aqUtils.Delay(5000,"Working Estimate")
ValidationUtils.verify(true,true,"Working Estimate is Selected");
TextUtils.writeLog("Working Estimate is Selected"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var approve_bar ;
  PropArray = new Array("JavaClassName", "Index","ChildCount");
  ValuesArray = new Array("PTabItemPanel", "3","1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists)&&(obj[i_count].isVisible())){
    approve_bar = obj[i_count].SWTObject("TabControl", "");
    break;      
  }
}
Sys.HighlightObject(approve_bar);
approve_bar.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

//Clicking Maximize Icon
ImageRepository.ImageSet.Maximize.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
var AprveAction;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("TabControl", "5");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible)
  if(obj[i_count].text=="All Approval Actions"){
  Sys.HighlightObject(obj[i_count]);
  AprveAction = obj[i_count];
  break;
 }
}
Sys.HighlightObject(AprveAction);
    Sys.HighlightObject(AprveAction)
    AprveAction.Click();
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }



var ApproverTable;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
if(obj[i_count].Visible)
ApproverTable = obj[i_count];
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
WorkspaceUtils.waitForObj(ApproverTable);

for(var i=0;i<ApproverTable.getItemCount();i++){   
var approvers="";
if(ApproverTable.getItem(i).getText_2(8)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(false,true,"Level "+i+"Is not Approved");
}
}


TextUtils.writeLog("Working Estimate is APPROVED in all levels"); 
    
//Closing Sliding Pane
var ApprovalTableBar;


var ApprovalTableBar ;
  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("TabControl", "1","true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="PTabItemPanel") && (obj[i_count].Parent.Index==1)){
    ApprovalTableBar = obj[i_count];
    break;      
  }
}
Sys.HighlightObject(ApprovalTableBar);
Sys.HighlightObject(ApprovalTableBar);    
    
Sys.HighlightObject(ApprovalTableBar)
Log.Message(ApprovalTableBar.FullName)
ApprovalTableBar.Click(); 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ImageRepository.ImageSet.Forward.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

    }
    
}


function validatingWorkEstimate(){ 

  
  
  // Getting Client Currency for Calculating Tax
  var BarStat = true;
  var ChildCount = 0;
    var Add = [];
   var Parent = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
         for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
      if((PChild.isVisible()) && (PChild.ChildCount==1)){
      Add[ChildCount] = PChild;
      ChildCount++;

     }
     }      
     
      Parent = "";
     var pos = 1000;
     for(var i=0;i<Add.length;i++){ 
     if(Add[i].Height<pos){ 
       pos = Add[i].Height;
       Log.Message(pos)
       Parent = Add[i];
     }     
     } 
   Parent = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "");
   Log.Message(Parent.FullName);
   Sys.HighlightObject(Parent);
    ChildCount = 0;
    Add = [];
     for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
     Log.Message(PChild.Name);
     if((PChild.isVisible()) && (PChild.JavaClassName=="Composite") && (PChild.ChildCount==1)){
         Add[ChildCount] = PChild;
         ChildCount++;
     }
     }
     
     Parent = "";
     var pos = 1000;
     for(var i=0;i<Add.length;i++){ 
     if(Add[i].Height<pos){ 
       pos = Add[i].Height;
       Log.Message(pos)
       Parent = Add[i];
     }     
     } 
    Log.Message(Parent.FullName)
    Sys.HighlightObject(Parent);
    C_Currency = Parent.SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
    Log.Message(C_Currency.FullName)
    Sys.HighlightObject(C_Currency);


    Log.Message(C_Currency.FullName)
    WorkspaceUtils.waitForObj(C_Currency);
    C_Currency = C_Currency.getText().OleValue.toString().trim();;
  Log.Message(C_Currency);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }



  
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("TabControl", "6");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
   
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Full Budget"){
  Sys.HighlightObject(obj[i_count]);
  FullBudget = obj[i_count];
  break;
  }
}
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  aqUtils.Delay(5000,Indicator.Text)
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
var linestatus = false;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Grid = "";
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible){
  Sys.HighlightObject(obj[i_count]);
  Grid = obj[i_count];
  break;
  }
}


WorkspaceUtils.waitForObj(Grid);
Estimate = [];
workEstimate = [];
var j=0;

for(var i=0;i<Grid.getItemCount();i++){ 
  var workcode = Grid.getItem(i).getText_2(0).OleValue.toString().trim();
  var description = Grid.getItem(i).getText_2(3).OleValue.toString().trim();
  var quantity = Grid.getItem(i).getText_2(6).OleValue.toString().trim();
  var costBase = Grid.getItem(i).getText_2(7).OleValue.toString().trim();
  var billingPrice = Grid.getItem(i).getText_2(9).OleValue.toString().trim();
  if((workcode!="")||(description!="")||(quantity!="")||(costBase!="")||(billingPrice!="")){ 
   workEstimate[j] = workcode+"*"+description+"*"+quantity+"*"+costBase+"*"+billingPrice+"*";
  if(EnvParams.Country.toUpperCase()=="INDIA"){ 
  var Ohsn = Grid.getItem(i).getText_2(12).OleValue.toString().trim();
  var Ihsn = Grid.getItem(i).getText_2(13).OleValue.toString().trim();
  workEstimate[j] = workEstimate[j]+Ohsn+"*"+Ihsn+"*";
  }
   j++;
  }
}

}



// Transfering Estimate to Quote
function transferToQuote(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}


  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var tranQuote;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Transfer to Quote"){
  Sys.HighlightObject(obj[i_count]);
  tranQuote = obj[i_count];
  break;
  }
}

WorkspaceUtils.waitForObj(tranQuote);
tranQuote.HoverMouse();
ReportUtils.logStep_Screenshot("");
tranQuote.Click();
aqUtils.Delay(5000, "Jobs - Budget");
    
p = eval(WorkspaceUtils.Sys_Maconomy_Parent);
w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim(), 2000);
if (w.Exists)
{ 
var OK = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}
     

aqUtils.Delay(10000, "Jobs - Budget");
    
p = eval(WorkspaceUtils.Sys_Maconomy_Parent);
w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim(), 2000);
if (w.Exists)
{ 
var OK = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}

ValidationUtils.verify(true,true,"Transfer to Quote is Clicked");
TextUtils.writeLog("Transfer to Quote is Clicked"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}



  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var quoteTab;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Quote"){
  Sys.HighlightObject(obj[i_count]);
  quoteTab = obj[i_count];
  break;
  }
}

WorkspaceUtils.waitForObj(quoteTab);
quoteTab.HoverMouse();
WorkspaceUtils.waitForObj(quoteTab);
quoteTab.Click();
aqUtils.Delay(3000, Indicator.Text);
     
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var specification = "";
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible){
  Sys.HighlightObject(obj[i_count]);
  specification = obj[i_count];
  break;
  }
}


  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "1", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var McGroupWidget = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==2))
McGroupWidget = obj[i_count];
}
Sys.HighlightObject(McGroupWidget)

Log.Message(McGroupWidget.Parent.Name)
var newQuote = McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
var EffQuotePrice = McGroupWidget.SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);


  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "2", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var McGroupWidget = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==1))
McGroupWidget = obj[i_count];
}
Sys.HighlightObject(McGroupWidget)
var Q_revision = McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);


newQuote = Q_revision.getText().OleValue.toString().trim();
EffQuotePrice = EffQuotePrice.getText().OleValue.toString().trim();
Q_revision = Q_revision.getText().OleValue.toString().trim();
var QuoteMPL = "QuoteMPL";


var q = 0;
QuoteDetails = [];
Log.Message(specification.getItemCount())
for(var i=0;i<specification.getItemCount();i++){ 
Log.Message("i: "+i);
  var Q_Desp = specification.getItem(i).getText_2(0).OleValue.toString().trim();
  if(Q_Desp!=""){
  var Q_Qty = specification.getItem(i).getText_2(1).OleValue.toString().trim();
  var Q_Billing = specification.getItem(i).getText_2(2).OleValue.toString().trim();
  var Q_BillingTotoal = specification.getItem(i).getText_2(3).OleValue.toString().trim();
  var Q_Tax1 = specification.getItem(i).getText_2(4).OleValue.toString().trim();
  var Q_Tax2 = specification.getItem(i).getText_2(5).OleValue.toString().trim();
  var Q_Tax1currency = specification.getItem(i).getText_2(9).OleValue.toString().trim();
  var Q_Tax2currency = specification.getItem(i).getText_2(10).OleValue.toString().trim();
  var Q_total = parseFloat(Q_BillingTotoal.replace(/,/g, ''))+ parseFloat(Q_Tax1currency.replace(/,/g, '')) + parseFloat(Q_Tax2currency.replace(/,/g, ''));
  QuoteDetails[q] = Q_Desp+"*"+Q_Qty+"*"+Q_Billing+"*"+Q_BillingTotoal+"*"+Q_Tax1+"*"+Q_Tax2+"*"+Q_Tax1currency+"*"+Q_Tax2currency+"*"+Q_total.toFixed(2)+"*";
  Log.Message(QuoteDetails[q]);
  q++;
  ExcelUtils.setExcelName(workBook,QuoteMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,QuoteMPL,Q_Desp);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,QuoteMPL,Q_Qty);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,QuoteMPL,Q_Billing);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TotalBilling_"+q,QuoteMPL,Q_BillingTotoal);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1_"+q,QuoteMPL,Q_Tax1);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2_"+q,QuoteMPL,Q_Tax2);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1currency_"+q,QuoteMPL,Q_Tax1currency);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2currency_"+q,QuoteMPL,Q_Tax2currency);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Total_"+q,QuoteMPL,Q_total);

  }
  }
Log.Message(q)
for(var i=q;i<11;i++){ 
  ExcelUtils.setExcelName(workBook,QuoteMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TotalBilling_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax1currency_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Tax2currency_"+q,QuoteMPL,"");
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Total_"+q,QuoteMPL,"");
}

ExcelUtils.setExcelName(workBook,QuoteMPL, true);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quote Revision",QuoteMPL,Q_revision);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quote Total",QuoteMPL,EffQuotePrice);



  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var printdraft = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Print Draft Quote"){
  Sys.HighlightObject(obj[i_count]);
  printdraft = obj[i_count];
  break;
 }
}

Log.Message(printdraft.FullName)
Sys.HighlightObject(printdraft);
WorkspaceUtils.waitForObj(printdraft);    
printdraft.HoverMouse();
ReportUtils.logStep_Screenshot("");
printdraft.Click();
TextUtils.writeLog("Print Draft Quote is Clicked and saved"); 
aqUtils.Delay(5000, Indicator.Text);
WorkspaceUtils.savePDF_localDirectory("PDF Print Draft Quote","Print Job Quote");




if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var submit = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Submit"){
  Sys.HighlightObject(obj[i_count]);
  submit = obj[i_count];
  break;
 }
}

Log.Message(submit.FullName)
Sys.HighlightObject(submit);
WorkspaceUtils.waitForObj(submit); 
submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
submit.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
aqUtils.Delay(5000, Indicator.Text);
while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    aqUtils.Delay(500, "Job Quote is loading");
} 

ValidationUtils.verify(true,true,"Quote has Submitted");
TextUtils.writeLog("Quote has Submitted"); 

  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("McPaneGui$10", "true");
p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
obj = p.FindAll(PropArray, ValuesArray, 1000);
var Page = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if(obj[i_count].Exists)
Page = obj[i_count];
}
Sys.HighlightObject(Page)
Page.Click();
Page.MouseWheel(-100);



  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "2", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var McGroupWidget = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==3))
McGroupWidget = obj[i_count];
}
Sys.HighlightObject(McGroupWidget)

Log.Message(McGroupWidget.Parent.Name)
//var submittedby = McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
var submittedby = McGroupWidget.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 2)
Sys.HighlightObject(submittedby);
submittedby.Click();
ValidationUtils.verify(true,true,"Quote is Submitted by :"+ submittedby.getText());
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
} 

if(EnvParams.Language.toUpperCase()=="SPANISH"){
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.Spanish.Close_Workspace.Exists()){ 
//ImageRepository.Spanish.Close_Workspace.Click();
//}
aqUtils.Delay(2000, Indicator.Text);
}


var approve_bar ;
  PropArray = new Array("JavaClassName", "Index","ChildCount");
  ValuesArray = new Array("PTabItemPanel", "3","1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists)&&(obj[i_count].isVisible())){
    approve_bar = obj[i_count].SWTObject("TabControl", "");
    break;      
  }
}
Sys.HighlightObject(approve_bar);
approve_bar.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 

var AprveAction;
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible)
  if(obj[i_count].text=="All Approval Actions"){
  Sys.HighlightObject(obj[i_count]);
  AprveAction = obj[i_count];
  break;
 }
}
Sys.HighlightObject(AprveAction);
    Sys.HighlightObject(AprveAction)
    AprveAction.Click();
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 

var Approval_table;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
if(obj[i_count].Visible)
Approval_table = obj[i_count];
}

    Log.Message(Approval_table.FullName)
    Sys.HighlightObject(Approval_table);
    Sys.HighlightObject(Approval_table)
    
    Project_manager = eval(Maconomy_ParentAddress).WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
    for(var z=0;z<Approval_table.getItemCount();z++){ 
       approvers="";
//       approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
//       Approve_Level[y] = comapany+"*"+jobNumber+"*"+approvers;
//       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
//       y++;
       
       if(Approval_table.getItem(z).getText_2(8)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
        
      //Self Approve is Disabled. So finding different Approver
        var mainApprover = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim();
        var substitur = Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+jobNumber+"*"+ temp;
      Log.Message("Approver level :" +z+ ": " +approvers);
      Log.Message(y);
      Log.Message(Approve_Level.length);
      Approve_Level[y] = approvers;
      y++;
      }
      
    }



}


//Converting Job Status from Quote to Order
function convertToOrder(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var home = ""
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Home"){
  Sys.HighlightObject(obj[i_count]);
  home = obj[i_count];
  break;
 }
}

Log.Message(home.FullName)
Sys.HighlightObject(home);
WorkspaceUtils.waitForObj(home); 
home.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Information = ""
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Information"){
  Sys.HighlightObject(obj[i_count]);
  Information = obj[i_count];
  break;
 }
}

Log.Message(Information.FullName)
Sys.HighlightObject(Information);
WorkspaceUtils.waitForObj(Information); 
Information.Click();
 TextUtils.writeLog("Navigated to Home"); 
 TextUtils.writeLog("Navigated to Information"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}



  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var convertToOrder = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Convert to Order"){
  Sys.HighlightObject(obj[i_count]);
  convertToOrder = obj[i_count];
  break;
 }
}

Log.Message(convertToOrder.FullName)
Sys.HighlightObject(convertToOrder);


WorkspaceUtils.waitForObj(convertToOrder);    
Sys.HighlightObject(convertToOrder)
convertToOrder.HoverMouse();
ReportUtils.logStep_Screenshot("");
convertToOrder.Click();
aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ValidationUtils.verify(true,true,"Convert To Order is Clicked");
TextUtils.writeLog("Convert To Order is Clicked"); 


  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "2", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var McGroupWidget = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==3))
McGroupWidget = obj[i_count];
}
Sys.HighlightObject(McGroupWidget)

Log.Message(McGroupWidget.Parent.Name)
var status = McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(status);

var i=0;
while ((status.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Order").OleValue.toString().trim())==-1)&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  status.Refresh();
}   
  
if(status.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Order").OleValue.toString().trim())!=-1){
Sys.HighlightObject(status);
aqUtils.Delay(3000, Indicator.Text);
Log.Message("Job Status :"+status.getText())
ReportUtils.logStep("INFO","Job Status :"+status.getText())  ;
TextUtils.writeLog("Job Status :"+status.getText()); 
}else{ 
TextUtils.writeLog("Convert to Order is Cliecked But its Not Reflected"); 
ValidationUtils.verify(true,false,"Convert to Order is Cliecked But its Not Reflected")
}
   

  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Budget;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Budgeting"){
  Sys.HighlightObject(obj[i_count]);
  Budget = obj[i_count];
  break;
  }
}
WorkspaceUtils.waitForObj(Budget); 
Budget.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(3000, Indicator.Text);


  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var quoteTab;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Quote"){
  Sys.HighlightObject(obj[i_count]);
  quoteTab = obj[i_count];
  break;
  }
}

WorkspaceUtils.waitForObj(quoteTab);
quoteTab.HoverMouse();
WorkspaceUtils.waitForObj(quoteTab);
quoteTab.Click();
aqUtils.Delay(3000, Indicator.Text);
     
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var Actions = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"GroupToolItemControl","Actions");
Sys.HighlightObject(Actions);
Actions.Click();

  Delay(5000);
  Actions.PopupMenu.Click("Print Order Confirmation");
     TextUtils.writeLog("Print Order Confirmation is Clicked");
    aqUtils.Delay(5000, Indicator.Text);
    aqUtils.Delay(3000, Indicator.Text);
WorkspaceUtils.savePDF_localDirectory("PDF Print Order Confirmation","Print Job Order Confirmation");


//  PropArray = new Array("JavaClassName", "Visible");
//  ValuesArray = new Array("SingleToolItemControl", "true");
//  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  var PrintOrderConfirm = "";
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if(obj[i_count].toolTipText=="Print Order Confirmation"){
//  Sys.HighlightObject(obj[i_count]);
//  PrintOrderConfirm = obj[i_count];
//  break;
// }
//}
//
//Log.Message(PrintOrderConfirm.FullName)
//Sys.HighlightObject(PrintOrderConfirm);
//WorkspaceUtils.waitForObj(PrintOrderConfirm);     
//Sys.HighlightObject(PrintOrderConfirm);
//PrintOrderConfirm.HoverMouse();
//ReportUtils.logStep_Screenshot(""); 
//PrintOrderConfirm.Click();
//TextUtils.writeLog("Print Order Confirmation is Clicked");
//aqUtils.Delay(3000, Indicator.Text);
//WorkspaceUtils.savePDF_localDirectory("PDF Print Order Confirmation","Print Job Order Confirmation");



}


function clientApprovedEsimate(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var budget;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Budget"){
  Sys.HighlightObject(obj[i_count]);
  budget = obj[i_count];
  break;
  }
}
WorkspaceUtils.waitForObj(budget); 
Sys.HighlightObject(budget);
budget.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
TextUtils.writeLog("Navigated to Budget from Budgeting");



//var revision = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.Composite2.PTabFolder.SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
//var revision = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget

 var revision;   
      PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "1", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var McGroupWidget = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].ChildCount>=8))
McGroupWidget = obj[i_count];
}
    Sys.HighlightObject(McGroupWidget);
var revision = McGroupWidget.SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
    
Sys.HighlightObject(revision);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

//    var ChildCount = 0;
//    var Add = [];
//    var Parent = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//    for(var i=0;i<Parent.ChildCount;i++){ 
//    var PChild = Parent.Child(i);
//    if((PChild.isVisible()) && (PChild.ChildCount==1)){
//    Add[ChildCount] = PChild;
//    ChildCount++;
//
//    }
//    }      
//     
//      Parent = "";
//     var pos = 1000;
//     for(var i=0;i<Add.length;i++){ 
//     if(Add[i].Height<pos){ 
//       pos = Add[i].Height;
//       Log.Message(pos)
//       Parent = Add[i];
//     }     
//     } 
//   Parent = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "");
//   Log.Message(Parent.FullName);
//   Sys.HighlightObject(Parent);
//    ChildCount = 0;
//    Add = [];
//     for(var i=0;i<Parent.ChildCount;i++){ 
//     var PChild = Parent.Child(i);
//     Log.Message(PChild.Name);
//     if((PChild.isVisible()) && (PChild.JavaClassName=="Composite") && (PChild.ChildCount==1)){
//         Add[ChildCount] = PChild;
//         ChildCount++;
//     }
//     }
//     
//     Parent = "";
//     var pos = 1000;
//     for(var i=0;i<Add.length;i++){ 
//     if(Add[i].Height<pos){ 
//       pos = Add[i].Height;
//       Log.Message(pos)
//       Parent = Add[i];
//     }     
//     } 
//    Log.Message(Parent.FullName)
//    Sys.HighlightObject(Parent);
//    var show_budget = Parent.SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

var show_budget = McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);

Log.Message(show_budget.FullName)
WorkspaceUtils.waitForObj(show_budget);
Sys.HighlightObject(show_budget);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
show_budget.Keys("Client Approved Estimate"); 
aqUtils.Delay(4000, Indicator.Text);
TextUtils.writeLog("Client Approved Estimate is Selected");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

////if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.isVisible())
//var copy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 11);
////else{ 
////var copy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 10);
////Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 11)
////}



  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var copy;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy").OleValue.toString().trim()){
  Sys.HighlightObject(obj[i_count]);
  copy = obj[i_count];
  break;
  }
}

WorkspaceUtils.waitForObj(copy);
WorkspaceUtils.waitForObj(copy);
Sys.HighlightObject(copy)
copy.HoverMouse();
ReportUtils.logStep_Screenshot(""); 
copy.Click();
TextUtils.writeLog("Copy Button is Clicked");

//if(copy.text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy").OleValue.toString().trim()){
//copy.HoverMouse();
//ReportUtils.logStep_Screenshot(""); 
//copy.Click();
//TextUtils.writeLog("Copy Button is Clicked");
//}else{
//
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.isVisible())
//copy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.SingleToolItemControl5;
//else
//copy = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 12);
//WorkspaceUtils.waitForObj(copy);
//copy.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//copy.Click();
//TextUtils.writeLog("Copy Button is Clicked");
//}

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim()
  
//////    var Job = Aliases.Maconomy.Shell3.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
    var Job = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
        Job.Click();
    WorkspaceUtils.waitForObj(Job);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    if(ImageRepository.ImageSet.Img_search.Exists()){ 
      
    }
    Job.Click();
    if(Job.getText()!=jobNumber){
    WorkspaceUtils.SearchByValues_all_Col_2(Job,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim())
    Job
    }
    var copy_Budget = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2)
    WorkspaceUtils.waitForObj(copy_Budget);
    Sys.HighlightObject(copy_Budget);
    copy_Budget.Keys("Working Estimate");
    aqUtils.Delay(1000, Indicator.Text);
    
    var copy_RevesionNo = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2)
    Sys.HighlightObject(copy_RevesionNo);
    copy_RevesionNo.setText(RevisionNo);
    
//    var copy_button = Aliases.Maconomy.Shell3.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy").OleValue.toString().trim());
    var copy_button = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy").OleValue.toString().trim());
    WorkspaceUtils.waitForObj(copy_Budget);
    Sys.HighlightObject(copy_button);
    copy_button.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    copy_button.Click();
    
    aqUtils.Delay(6000, Indicator.Text);
    TextUtils.writeLog("Copying Working Estimate from Job Number :"+jobNumber+" to Client Approved Estimate" );
    p = eval(WorkspaceUtils.Sys_Maconomy_Parent);
    w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job Budgets Card API").OleValue.toString().trim(), 2000);
    if (w.Exists)
    { 
    var ApiButton = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job Budgets Card API").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
    ApiButton.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    ApiButton.Click();
    }
    
}


function validatingclientEstimate(){ 
var FullBudget = "";
var kk= 0;
//var mainRoot = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//var line = false;
//
//
//  if(!line){
//  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
//  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
//  if(tempName.indexOf("2")!=-1){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
//  FullBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
//  line = true;
//  break;
//  }
//  }
//  }
//  }
//  }
//  
//  
//  if(!line){
//  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
//  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
//  if(tempName.indexOf("3")!=-1){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
//  FullBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
//  line = true;
//  break;
//  }
//  }
//  }
//  }
//  }
//  
//  
//  if(!line){
//  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
//  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
//  if(tempName.indexOf("4")!=-1){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
//  FullBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
//  line = true;
//  break;
//  }
//  }
//  }
//  }
//  }
//  
//  
//  if(!line){
//  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
//  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
//  if(tempName.indexOf("5")!=-1){
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
//  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
//  FullBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
//  line = true;
//  break;
//  }
//  }
//  }
//  }
//  }
//  WorkspaceUtils.waitForObj(FullBudget);
//  Sys.HighlightObject(FullBudget)  ;
//  FullBudget.Click();


  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("TabControl", "6");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
   
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text=="Full Budget"){
  Sys.HighlightObject(obj[i_count]);
  FullBudget = obj[i_count];
  break;
  }
}
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

 var Grid = Clientgrid();  
 WorkspaceUtils.waitForObj(Grid);  
 WorkspaceUtils.waitForObj(Grid);
 WorkspaceUtils.waitForObj(Grid);     

clientEstimate = [];
var j=0;
//Log.Message(Grid.getItemCount())
for(var i=0;i<Grid.getItemCount();i++){ 
  var workcode = Grid.getItem(i).getText_2(0).OleValue.toString().trim();
  var description = Grid.getItem(i).getText_2(3).OleValue.toString().trim();
  var quantity = Grid.getItem(i).getText_2(6).OleValue.toString().trim();
  var costBase = Grid.getItem(i).getText_2(7).OleValue.toString().trim();
  var billingPrice = Grid.getItem(i).getText_2(9).OleValue.toString().trim();
  clientEstimate[j] = workcode+"*"+description+"*"+quantity+"*"+costBase+"*"+billingPrice+"*";
  
  if(EnvParams.Country.toUpperCase()=="INDIA"){ 
  var Ohsn = Grid.getItem(i).getText_2(12).OleValue.toString().trim();
  var Ihsn = Grid.getItem(i).getText_2(13).OleValue.toString().trim();
  clientEstimate[j] = clientEstimate[j]+Ohsn+"*"+Ihsn+"*";
  }
   j++;
}

//finding Duplicate Rows
var Duplicate = [];
var Unique = [];
var z=0;
var y=0;
//Log.Message(clientEstimate.length)
for(var i=0;i<clientEstimate.length;i++){
if(clientEstimate[i]=="**0.00*0.00*0.00*"){ 
 Duplicate[y]= "**0.00*0.00*0.00*";
 y++;
}
else if(Unique.indexOf(clientEstimate[i])==-1){
Unique[z]=clientEstimate[i];
//Log.Message("Unique[z] :"+Unique[z])
z++;
}
else{
Duplicate[y]=clientEstimate[i];
Log.Message("Duplicate[y] :"+Duplicate[y])
y++;
}
  }
  
var Adding = [];
z=0;
for(var i=0;i<workEstimate.length;i++){
  var temp = false;
  for(var j=0;j<clientEstimate.length;j++){ 
    if(workEstimate[i]==clientEstimate[j])
    temp = true;
    }
  if(!temp){ 
    Adding[z] = workEstimate[i];
    z++;
  }
  
  
}

//Removing Duplicate Lines


var deleteBudget = "";
var itemCount = Grid.getItemCount();
for(var i=0;i<itemCount;i++){ 
var workcode = Grid.getItem(i).getText_2(0).OleValue.toString().trim();
var description = Grid.getItem(i).getText_2(3).OleValue.toString().trim();
var quantity = Grid.getItem(i).getText_2(6).OleValue.toString().trim();
var costBase = Grid.getItem(i).getText_2(7).OleValue.toString().trim();
var billingPrice = Grid.getItem(i).getText_2(9).OleValue.toString().trim();
  if(EnvParams.Country.toUpperCase()=="INDIA"){ 
  var Ohsn = Grid.getItem(i).getText_2(12).OleValue.toString().trim();
  var Ihsn = Grid.getItem(i).getText_2(13).OleValue.toString().trim();
}

var temp = workcode+"*"+description+"*"+quantity+"*"+costBase+"*"+billingPrice+"*";
  
if(EnvParams.Country.toUpperCase()=="INDIA"){
  temp = temp+Ohsn+"*"+Ihsn+"*";
  }
  
var matchStatus = false;
if(Duplicate.length!=0){ 
ReportUtils.logStep("WARNING","Some Duplicate Budget lines or Extra Budget lines are there");
}
for(var j=0;j<Duplicate.length;j++){

if(temp==Duplicate[j]){
matchStatus = true;
var deleteBudget = "";
//var kk= 0;
var mainRoot = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
var line = false;
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("2")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  deleteBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("3")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  deleteBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);;
  line = true;
  break;
  }
  }
  }
  }
  } 
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("4")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  deleteBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("5")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  deleteBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("7")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  deleteBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);;
  line = true;
  break;
  }
  }
  }
  }
  }

  
  Sys.HighlightObject(deleteBudget)  ;
  deleteBudget.HoverMouse();
ReportUtils.logStep_Screenshot("");
if(deleteBudget.toolTipText.OleValue.toString().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete Job Budget Line").OleValue.toString().trim())!=-1){ 
deleteBudget.Click();
var OK = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.HoverMouse();
ReportUtils.logStep_Screenshot("");
OK.Click();

  aqUtils.Delay(4000, "Reading Table data after Delete");
  i=i-1;
  itemCount = Grid.getItemCount();
  Duplicate[j]="";
  break;
  }
  }
  }
  if(!matchStatus){ 

  itemCount = Grid.getItemCount();

  if(i<(Grid.getItemCount()-2)){
    Grid.Keys("[Down]");
    aqUtils.Delay(1000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    }
    aqUtils.Delay(1000, Indicator.Text);
  }
  }
  eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  
var addedlines = false; 

if(Adding.length!=0){ 
ReportUtils.logStep("WARNING","Some Budget lines are missed while copying from Work Estimate");
ReportUtils.logStep("INFO","Adding Budget lines which are missed while copying from Work Estimate");
}
 for(var i=0;i<Adding.length;i++){
var temp = Adding[i].split("*");
var wCodeID = temp[0];
var Desp = temp[1];
var Qly = temp[2];
var CostBase = temp[3];
var bill = temp[4];

var Country = EnvParams.Country;
if(Country.toUpperCase()=="INDIA")
{
var OutwardHSN = temp[5]   
var InwardHSN = temp[6] 
}

if((wCodeID!="")&&(wCodeID!=null)){
var AddBudget = "";
var mainRoot = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
var line = false;
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("2")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  AddBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("3")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  AddBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);;
  line = true;
  break;
  }
  }
  }
  }
  } 
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("4")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  AddBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("5")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  AddBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("7")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  AddBudget = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);;
  line = true;
  break;
  }
  }
  }
  }
  }


Sys.HighlightObject(AddBudget)  ;
AddBudget.HoverMouse();
ReportUtils.logStep_Screenshot("");
AddBudget.Click(); 
  
  eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    
//    Delay(3000);
//-----Work Code Selection---------    
//var workcode = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McValuePickerWidget", "");
var workcode = Clientgrid().SWTObject("McValuePickerWidget", "");;
WorkspaceUtils.waitForObj(workcode);
if(wCodeID!=""){
addedlines = true;
  workcode.Click();
  WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"Work Code :"+wCodeID);
         }else{ 
    ValidationUtils.verify(false,true,"WorkCode Needed to create JobBudget");
  }
aqUtils.Delay(2000, Indicator.Text);  
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

//-----Internal Description---------
linestatus = false;
//var Internal_Description = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;;
//var Internal_Description = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McTextWidget", "", 2);
var Internal_Description = Clientgrid().SWTObject("McTextWidget", "", 4);

    Sys.HighlightObject(Internal_Description);
    Internal_Description.Click();
    if(Desp!=""){
    Internal_Description.setText(Desp);
    ValidationUtils.verify(true,true,"Internal Description is Entered");
    }
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

//-----Employee Categories if required---------
         
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);

    
//-----Employee Number if required---------    
         
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
    
//-----Quantity---------

//var Quality = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;;
//var Quality = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McTextWidget", "", 2);
var Quality = Clientgrid().SWTObject("McTextWidget", "", 2);
linestatus = false;


    Sys.HighlightObject(Quality);
    Quality.Click();
    if(Qly!=""){
    Quality.setText(Qly);
    ValidationUtils.verify(true,true,"Quality is Entered");
    }

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(2000, Indicator.Text);
    
    
//-----Cost Base Only for Amount---------
//    if(wCodeID.indexOf("T")==-1){
      
//var Cost_base = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;;
//var Cost_base = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McTextWidget", "", 2);
var Cost_base = Clientgrid().SWTObject("McTextWidget", "", 2);
linestatus = false;



    Sys.HighlightObject(Cost_base);
    Cost_base.Click();
    if(CostBase!=""){    
    Cost_base.setText(CostBase);
    ValidationUtils.verify(true,true,"Cost is Entered");
    }
//    }
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    
//-----Cost Base Only for Time---------
//    if(wCodeID.indexOf("T")>-1){
      
//var Billing_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;;
//var Billing_Price = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McTextWidget", "", 2);
var Billing_Price = Clientgrid().SWTObject("McTextWidget", "", 2);
linestatus = false;


    Sys.HighlightObject(Billing_Price);
    Billing_Price.Click();
    if(bill!=""){      
    Billing_Price.setText(bill);
    ValidationUtils.verify(true,true,"Billing Price is Entered is Entered");
    }
    
var Country = EnvParams.Country;
if(Country.toUpperCase()=="INDIA")
{

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    
  
    
//-----Outward HSN for INDIA---------
var Outward_HSN;
 Outward_HSN = Clientgrid().SWTObject("McValuePickerWidget", "", 1);
linestatus = false;
    
Sys.HighlightObject(Outward_HSN);
if(OutwardHSN!=""){
Outward_HSN.Click();
WorkspaceUtils.SearchByValue(Outward_HSN,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 8").OleValue.toString().trim(),OutwardHSN,"Outward HSN");
     }else{ 
ValidationUtils.verify(false,true,"Outward_HSN Needed to create JobBudget");
}

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    
    
//-----Invard HSN for INDIA---------
var Invard_HSN;
 Invard_HSN = Clientgrid().SWTObject("McValuePickerWidget", "", 1);
linestatus = false;
    
Sys.HighlightObject(Invard_HSN);
if(InwardHSN!=""){
Invard_HSN.Click();
WorkspaceUtils.SearchByValue(Invard_HSN,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 9").OleValue.toString().trim(),InwardHSN,"Inward HSN");
     }else{ 
ValidationUtils.verify(false,true,"Inward HSN Needed to create JobBudget");
}
  
}

linestatus = false;
var Save = "";


var kk= 0;
var mainRoot = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
var line = false;
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("2")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2)
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("3")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }

if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("4")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("5")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("7")!=-1){
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }


    Sys.HighlightObject(Save);
    Save.HoverMouse();
ReportUtils.logStep_Screenshot("");
    Save.Click();

} 
}
/*

var tableGrid = Clientgrid();
var TotalBudget = 0.00;
for(var ti=0;ti<tableGrid.getItemCount();ti++){ 
  if(tableGrid.getItem(ti).getText_2(10).OleValue.toString().trim()!=""){
      var Tax_code1 = tableGrid.getItem(ti).getText_2(13).OleValue.toString().trim();
      var Tax_code2 = tableGrid.getItem(ti).getText_2(14).OleValue.toString().trim();
    if((Tax_code1!="")||(Tax_code2!="")){
    var total_Billing_Price_Currency = tableGrid.getItem(ti).getText_2(10).OleValue.toString().trim();
    total_Billing_Price_Currency = total_Billing_Price_Currency.replace(/,/g, '');
    TotalBudget = parseFloat(TotalBudget.toString()) + parseFloat(total_Billing_Price_Currency.toString());
    }else{ 
      ValidationUtils.verify(false,true,"Tax Code 1 and Tax Code 2 is not Populated in full Budget");
    }
    }
}
Log.Message(TotalBudget);
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.ChildCount==1)
var total = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();
            
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.ChildCount==1)
//var total = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget.getText().OleValue.toString().trim();

total = total.replace(/,/g, '');
var tBPC = parseFloat(total.toString());
Log.Message(tBPC)
if(tBPC.toString().trim()==TotalBudget.toString().trim())
ValidationUtils.verify(true,true,"Budget Currency is verified");
else
ValidationUtils.verify(false,true,"Budget Currency is Not Matched ");        
TextUtils.writeLog("Client Approved Estimate is validated");
TextUtils.writeLog("Total Budget Currency is Matched");

*/
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  
  
var Submit = "";   
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Submit"){
  Sys.HighlightObject(obj[i_count]);
  Submit = obj[i_count];
  break;
 }
}

Log.Message(Submit.FullName)
WorkspaceUtils.waitForObj(Submit);
ReportUtils.logStep_Screenshot("");
Submit.Click();

ReportUtils.logStep_Screenshot("");          
var Add_Visible8 = true;

ValidationUtils.verify(true,true,"Created Budget is Submitted");
TextUtils.writeLog("Client Approved Estimate is Submitted");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var IndiaSpec = "";
 
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.isVisible())
//IndiaSpec = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.PTabItemPanel.TabControl
//else if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.isVisible())
//IndiaSpec = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.PTabItemPanel.TabControl;
//
//IndiaSpec.Click();

var approve_bar ;
  PropArray = new Array("JavaClassName", "Index","ChildCount");
  ValuesArray = new Array("PTabItemPanel", "3","1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists)&&(obj[i_count].isVisible())){
    approve_bar = obj[i_count].SWTObject("TabControl", "");
    break;      
  }
}
Sys.HighlightObject(approve_bar);
approve_bar.Click();


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ImageRepository.ImageSet.Maximize.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var AprveAction;
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible)
  if(obj[i_count].text=="All Approval Actions"){
  Sys.HighlightObject(obj[i_count]);
  AprveAction = obj[i_count];
  break;
 }
}
Sys.HighlightObject(AprveAction);
    Sys.HighlightObject(AprveAction)
    AprveAction.Click();
    
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
   
var y=0;
var Approval_table;
  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
if(obj[i_count].Visible)
Approval_table = obj[i_count];
}

    Log.Message(Approval_table.FullName)
    Sys.HighlightObject(Approval_table);
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
Sys.HighlightObject(Approval_table)
    
    for(var z=0;z<Approval_table.getItemCount();z++){ 
       approvers="";
       approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
       Approve_Level[y] = comapany+"*"+jobNumber+"*"+approvers;
       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
       y++;
    }
    
    
var ApprovalTableBar ;
  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("TabControl", "1","true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="PTabItemPanel") && (obj[i_count].Parent.Index==1)){
    ApprovalTableBar = obj[i_count];
    break;      
  }
}
Sys.HighlightObject(ApprovalTableBar);
Sys.HighlightObject(ApprovalTableBar);    
    
Sys.HighlightObject(ApprovalTableBar)
Log.Message(ApprovalTableBar.FullName)
ApprovalTableBar.Click(); 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

ImageRepository.ImageSet.Forward.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


//CredentialLogin();
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);

var OpCo2 = ApproveInfo[0].split("*");
Project_manager = eval(Maconomy_ParentAddress).WndCaption;
Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);

//sheetName = "CreateQuote";
if(OpCo2[2]==Project_manager){
level = 1;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Approve = "";   
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Approve"){
  Sys.HighlightObject(obj[i_count]);
  Approve = obj[i_count];
  break;
 }
}
Sys.HighlightObject(Approve)
Sys.HighlightObject(Approve)
Approve.Click();
ReportUtils.logStep_Screenshot("");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
ReportUtils.logStep_Screenshot("");

var loginPer = eval(Maconomy_ParentAddress).WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);

  ValidationUtils.verify(true,true,"Client Approved Estimate is Approved by :"+loginPer)
  TextUtils.writeLog("Client Approved Estimate is Approved by :"+loginPer); 

}

if(ApproveInfo.length == 1){
  try{
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Client Approved Estimate_"+Jobnum,EnvParams.Opco,"Data Management",jobNumber);
}
catch(err){ 
  Log.Message(err);
}
}

}





// Finding Created Budget from To-Do's List
function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  
  var toDo = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
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


//Client_Managt =  eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "")

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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job Budget by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job Budget by Type from To-Dos List");
listPass = false; 
  }
}

if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job Budget by Type (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job Budget by Type (Substitute) from To-Dos List");
var listPass = false;   
  }
}  

if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job Budget").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job Budget from To-Dos List");
listPass = false; 
  }
}

if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job Budget (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job Budget (Substitute) from To-Dos List");
var listPass = false;   
  }
} 
  }

  
if(listPass){
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job Budget by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job Budget by Type from To-Dos List");
listPass = false; 
  }
}
}

if(listPass){
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Job Budget by Type (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job Budget by Type (Substitute) from To-Dos List");
var listPass = false;   
  }
} 
  }

  

}


function aprvBudget(ComId,JobNo,userNmae){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "");
waitForObj(table);
Sys.HighlightObject(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).Visible){

}else{ 
var showFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
waitForObj(table);
Sys.HighlightObject(showFilter);
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.Click();
}

//var allJobs = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
//allJobs.Click();
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
//var closeFilter = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//var table = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//var companyFilter = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;


 var closeFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   var companyFilter = eval(Maconomy_ParentAddress).
    SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
    SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).
    SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
    SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
    SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    
companyFilter.forceFocus();
companyFilter.setVisible(true);
companyFilter.ClickM();
table.Child(0).setText("^a[BS]");
table.Child(0).setText(ComId);
WorkspaceUtils.waitForObj(table)

table.Child(0).Keys("[Tab][Tab]");

var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    
job.ClickM();
table.Child(2).forceFocus();
table.Child(2).setVisible(true);
table.Child(2).setText("^a[BS]");
table.Child(2).setText(JobNo);
WorkspaceUtils.waitForObj(table)

aqUtils.Delay(3000, "Reading Data from table");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
if(table.getItem(v).getText_2(2).OleValue.toString().trim()==JobNo){ 
  flag=true;
  break;
}
else{ 
  table.Keys("[Down]");
}
}
ValidationUtils.verify(flag,true,"Job is listed for Approval");
TextUtils.writeLog("Created JobBudget is listed in Approval list");
if(table.getItemCount()>0){
closeFilter.HoverMouse();
ReportUtils.logStep_Screenshot("");
closeFilter.Click();

    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, "Waiting for Maconomy to load screen fully");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
  aqUtils.Delay(5000, Indicator.Text);
  
   var show_budget;   
      PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGroupWidget", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].ChildCount>=8))
show_budget = obj[i_count];
}
    Sys.HighlightObject(show_budget);
show_budget = show_budget.SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);
    WorkspaceUtils.waitForObj(show_budget);
    show_budget.HoverMouse();
    show_budget.HoverMouse();
    show_budget.HoverMouse();

    Sys.HighlightObject(show_budget);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    WorkspaceUtils.waitForObj(show_budget)
    Sys.HighlightObject(show_budget);
    show_budget.Keys("Client Approved Estimate");
    
    aqUtils.Delay(5000, "Client Approved Estimate");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(5000, "Client Approved Estimate");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }



var Approve = "";   
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Approve"){
  Sys.HighlightObject(obj[i_count]);
  Approve = obj[i_count];
  break;
 }
}
Sys.HighlightObject(Approve)
Sys.HighlightObject(Approve)
WorkspaceUtils.waitForObj(Approve);

if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot("");
Approve.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}




var loginPer = eval(Maconomy_ParentAddress).WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);

  ValidationUtils.verify(true,true,"Client Approved Estimate is Approved by :"+loginPer)
  TextUtils.writeLog("Client Approved Estimate is Approved by :"+loginPer); 

}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  TextUtils.writeLog("Approve Button Is Invisible"); 
  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
  TextUtils.writeLog(ComId+" - "+JobNo +" - Approver :"+userNmae); 
}
    }

    
    
if((ApproveInfo.length -1)== level){
try{
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Client Approved Estimate_"+Jobnum,EnvParams.Opco,"Data Management",jobNumber);
}catch(err){ 
  Log.Message(err);
}
}
}


function Clientgrid(){ 
  

     
       PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("McGrid", "2");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var grid = "";
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible){
  Sys.HighlightObject(obj[i_count]);
  grid = obj[i_count];
  break;
  }
}

//var Grid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

WorkspaceUtils.waitForObj(grid);
return grid;
}





function ApproveQuote(){ 
  
Log.Message(ApproveInfo.length)
Log.Message(level);
  // Approving Created Job Quote
for(var i=level;i<ApproveInfo.length;i++){
  level = i;
  
  

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
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
aqUtils.Delay(5000, Indicator.Text);

Log.Message("i :"+i);
Log.Message("temp[3] :"+temp[3])
ActionUtils.ToDos_Selection(Maconomy_ParentAddress, i, temp[3], null,null,"Approve Job Quote by Type (Substitute)",null)
ApproveDraftQuote(temp[0],temp[1],temp[2]);

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);
}


}



function ApproveDraftQuote(comID,JobNum,AprName){ 
  

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
                                         
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "");
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

if(eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
                                         
var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2); 
var firstCell = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
firstCell.Keys(comID);
firstCell.Keys("[Tab][Tab]");

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var JobNumber = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
JobNumber.Click();
JobNumber.Keys(JobNum);

var closefilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
aqUtils.Delay(3000, Indicator.Text); 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

for(var v=0;v<table.getItemCount();v++){ 
  WorkspaceUtils.waitForObj(table);
var flag=false;
  if((table.getItem(v).getText_2(0).OleValue.toString().trim()==comID)||
  (table.getItem(v).getText_2(2).OleValue.toString().trim()==JobNum)){ 


    flag=true; 
    table.Keys("[Down]");
    }else{ 
      table.Keys("[Down]");
    }
  }
    
    ValidationUtils.verify(flag,true,"Created Draft Quote is available in system");
    ReportUtils.logStep_Screenshot();  
    if(flag){ 
   closefilter.Click();


waitUntil_MaconomyScreen_loaded_Completely();
var Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve");
Sys.HighlightObject(Approve);
Approve.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


  PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "2", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var McGroupWidget = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==3))
McGroupWidget = obj[i_count];
}
Sys.HighlightObject(McGroupWidget)

Log.Message(McGroupWidget.Parent.Name)

var approveBy = McGroupWidget.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2)
Sys.HighlightObject(approveBy);
approveBy.Click();
ValidationUtils.verify(true,true,"Quote is Approved by :"+ approveBy.getText());

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var Actions = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"GroupToolItemControl","Actions");
Sys.HighlightObject(Actions);
Actions.Click();

  Delay(5000);
  Actions.PopupMenu.Click("Print Quote");
     TextUtils.writeLog("Print Quote is Clicked");
    aqUtils.Delay(5000, Indicator.Text);
    WorkspaceUtils.savePDF_localDirectory("PDF Quote","Print Job Quote");

}

}




//Validating Working Estimate is Approved or Not
function SelectJob(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  
  var allJobs = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  
  var table = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  
  var firstcell = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  
  var closeFilter = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
  
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
    
  ReportUtils.logStep("INFO", "Job is listed in table to create Quote");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job is available in maconommy to create Quote"); 
  closeFilter.Click();
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}
}