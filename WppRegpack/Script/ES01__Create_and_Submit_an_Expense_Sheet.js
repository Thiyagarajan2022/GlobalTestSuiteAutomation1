﻿//USEUNIT ActionUtils
//USEUNIT EnvParams
//USEUNIT EventHandler
//USEUNIT ExcelUtils
//USEUNIT ObjectUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/**
 * This script create Quote and Client Approved Estimate for Main Job
 * @author  : Muthu Kumar M
 * @version : 3.0
 * Created Date :02/10/2021
 * Modified Date(MM/DD/YYYY) : 01/06/2022
*/


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateExpense";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");
  
  
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var employeeNo,Description,desp,VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
var Language = "";



var Arrays = [];
var count = true;
var STIME = "";
var Description;
var jobNumber = "";
var Language = "";

//Main Function
function CreateExpense() {
TextUtils.writeLog("Create Purchase Order Started"); 
Indicator.PushText("waiting for window to open");


//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;


//Checking Login to execute Create Budget
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


Arrays = [];
count = true;
STIME = "";
Description;
jobNumber = "";

          STIME = WorkspaceUtils.StartTime();
          getDetails();
          goTo_TimeSheet();
          newExpenseSheet();
          gotoTimeExpenses();
          WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){

    ExcelUtils.setExcelName(workBook, sheetName, true);
    Description= ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
    if((Description==null)||(Description=="")){ 
    ValidationUtils.verify(false,true,"Description is Needed to Create a Expenses");
    }
    
    employeeNo= ExcelUtils.getColumnDatas("Employeeno",EnvParams.Opco)
    if((employeeNo==null)||(employeeNo=="")){ 
    ValidationUtils.verify(false,true,"Employee NO is Needed to Create a Expenses");
    }
    

sheetName ="CreateExpense";
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  sheetName ="CreateExpense";
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Expenses");
    
    
var CodeStatus = true;
var Country = EnvParams.Country;

 for(var i=1;i<=10;i++){
ExcelUtils.setExcelName(workBook, sheetName, true);
var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
var Curr = ExcelUtils.getColumnDatas("currency_"+i,EnvParams.Opco)
var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
var Amt = ExcelUtils.getColumnDatas("Amount_"+i,EnvParams.Opco)
var ExpRes =  ExcelUtils.getColumnDatas("Expense Reason_"+i,EnvParams.Opco)
var Vname = ExcelUtils.getColumnDatas("Vendor Name_"+i,EnvParams.Opco)
var GSTIN = ExcelUtils.getColumnDatas("GSTIN_"+i,EnvParams.Opco)
var InvoiceNo = ExcelUtils.getColumnDatas("Invoice No_"+i,EnvParams.Opco)
var InvoiceDate = ExcelUtils.getColumnDatas("Invoice Date_"+i,EnvParams.Opco)

if((wCodeID!="")&&(wCodeID!=null)){
  CodeStatus = false;
  if((Curr=="")||(Curr==null))
  ValidationUtils.verify(false,true,"currency_"+i+" is needed to Create Expenses");

//  if((Qly=="")||(Qly==null))
//  ValidationUtils.verify(false,true,"Quantity_"+i+" is needed to Create Expenses");
  
  if((Amt=="")||(Amt==null))
  ValidationUtils.verify(false,true,"Amount_"+i+" is needed to Create Expenses");
  
  if(Country.toUpperCase()=="INDIA"){ 
//  if((Vname=="")||(Vname==null))
//  ValidationUtils.verify(false,true,"Vendor Name_"+i+" is needed to Create Expenses");
  
  if((ExpRes=="")||(ExpRes==null))
  ValidationUtils.verify(false,true,"Expense Reason_"+i+" is needed to Create Expenses");
  }
  
}
}

if(CodeStatus)
ValidationUtils.verify(false,true,"WorkCode is needed to Create Expenses");

}



////------Label Validating Field-----////

function address(){
aqUtils.Delay(4000, Indicator.Text);
Sys.Process("Maconomy").Refresh();
var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McTextWidget", "", 1,60000).getText().OleValue.toString().trim();
if(employee!="Employee")
ValidationUtils.verify(false,true,"Employee field is missing in macanomy for the Expenses Creation");
else
ValidationUtils.verify(true,true,"Employee field is available in Macanomy for the Expenses Creation");

var description = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).WaitSWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
if(description!="Description")
ValidationUtils.verify(false,true,"Description field is missing in macanomy for the Expenses Creation");
else
ValidationUtils.verify(true,true,"Description field is available in Macanomy for the Expenses Creation");

var job = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").WaitSWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
if(job!="Job")
ValidationUtils.verify(fals,true,"Job field is missing in macanomy for the Expenses Creation");
else
ValidationUtils.verify(true,true,"Job field is available in Macanomy for the Expenses Creation");
}


// Navigating to Time & Expenses from Time & Expenses Menu
function goTo_TimeSheet(){

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_timesheet_from_workspace(); //Select Timesheet & Expense Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());

}




function newExpenseSheet(){ 
  ReportUtils.logStep("INFO", "Enter Expenses Details");
waitUntil_MaconomyScreen_loaded_Completely();
var ExpenseTab = getObjectAddress_JavaClasssName_Index_withTabText(Maconomy_ParentAddress,"TabControl",5,"Expenses");
  waitForObj(ExpenseTab)
  ReportUtils.logStep_Screenshot("");
  ExpenseTab.Click();

  waitUntil_MaconomyScreen_loaded_Completely();
  var AllExpenses = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"Button","All Expense Sheets");
  waitUntil_MaconomyScreen_loaded_Completely();
  
  Sys.Desktop.KeyDown(0x11)
  Sys.Desktop.KeyDown(0x46)
  aqUtils.Delay(3000, "Create Expenses Sheet");
  Sys.Desktop.KeyUp(0x11)
  Sys.Desktop.KeyUp(0x46)
  
var expenses = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl","New Expense Sheet (Ctrl+N)");


  Log.Message(expenses.FullName)
  waitForObj(expenses)
  WorkspaceUtils.waitForObj(expenses);
  ReportUtils.logStep_Screenshot("");
  expenses.Click();
  TextUtils.writeLog("Create New Expense Sheet is Clicked");
  
  waitUntil_MaconomyScreen_loaded_Completely();
  
  var Cancel = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(Cancel)

  var employee = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  WorkspaceUtils.waitForObj(employee)
  if(employee.getText()!=employeeNo){
  Sys.HighlightObject(employee);
  employee.HoverMouse();
  employee.Click();
     WorkspaceUtils.SearchByValue(employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),employeeNo,"Employee Number");
  }
  else{
  ValidationUtils.verify(true,true,"Employee Number is Exist in the Create Expenses");
  } 
  
  var descrip = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  descrip.HoverMouse();
  Sys.HighlightObject(descrip)
  WorkspaceUtils.waitForObj(descrip)
  desp = Description+" "+STIME;
  descrip.setText(desp); 
  
  var job = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
  job.HoverMouse();
  Sys.HighlightObject(job)
  job.HoverMouse();
  if(job.getText()!=jobNumber){
   job.Click();
   WorkspaceUtils.SearchByValues(job,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number");
  }
  else{ 
  ValidationUtils.verify(false,true,"Job Number is Exist in the Create Expenses");
  } 
  
  var Create = eval(WorkspaceUtils.Sys_Maconomy_Parent).SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim())
  WorkspaceUtils.waitForObj(descrip);
  ReportUtils.logStep_Screenshot(""); 
  Create.Click();
  TextUtils.writeLog("Expense Sheet is Created");
}



function gotoTimeExpenses(){
  
waitUntil_MaconomyScreen_loaded_Completely();
var ExpenseNumber = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McTextWidget",3);

WorkspaceUtils.waitForObj(ExpenseNumber);
ExpenseNumber = ExpenseNumber.getText().OleValue.toString().trim();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 

  }
  


var addedlines = false; 
 for(var i=1;i<=10;i++){
ExcelUtils.setExcelName(workBook, sheetName, true);
var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
var curr = ExcelUtils.getColumnDatas("currency_"+i,EnvParams.Opco)
var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
var Amt = ExcelUtils.getColumnDatas("Amount_"+i,EnvParams.Opco)
var Ereason =  ExcelUtils.getColumnDatas("Expense Reason_"+i,EnvParams.Opco)
var Vname = ExcelUtils.getColumnDatas("Vendor Name_"+i,EnvParams.Opco)
var GSTIN = ExcelUtils.getColumnDatas("GSTIN_"+i,EnvParams.Opco)
var I_no = ExcelUtils.getColumnDatas("Invoice No_"+i,EnvParams.Opco)
var I_Date = ExcelUtils.getColumnDatas("Invoice Date_"+i,EnvParams.Opco)

if((wCodeID!="")&&(wCodeID!=null)){
addedlines = true;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 
var addline = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl","Add Expense Sheet Line (Ctrl+M)");


WorkspaceUtils.waitForObj(addline);
ReportUtils.logStep_Screenshot();
addline.Click();

waitUntil_MaconomyScreen_loaded_Completely();

  
var EntryDate = getObjectAddress_withSingleProperty_Check(Maconomy_ParentAddress,"McDatePickerWidget");

WorkspaceUtils.waitForObj(EntryDate);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    
var workCode = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McValuePickerWidget",2,"McGrid"); 

workCode.Click();
WorkspaceUtils.SearchByValue(workCode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"Work Code :"+wCodeID);
Sys.HighlightObject(workCode);
var Wdes = workCode.getText();
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;

var WDesp = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McTextWidget",3,"McGrid"); 

WDesp.setText(Wdes);
Sys.HighlightObject(WDesp);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    
var Currency = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McPopupPickerWidget",4,"McGrid"); 

Currency.Keys(" ");
Currency.HoverMouse();
Sys.HighlightObject(Currency);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
if(curr!=""){
Currency.Click();
WorkspaceUtils.DropDownList(curr,"Currency")
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    
var UnitPrice = getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,"McTextWidget",3,"McGrid"); 

Sys.HighlightObject(UnitPrice);
UnitPrice.setText(Amt);

var save =  getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl", "Save Expense Sheet Line (Enter)");; 

WorkspaceUtils.waitForObj(save);
ReportUtils.logStep_Screenshot();
save.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

if(EnvParams.Country.toUpperCase()=="INDIA"){
Runner.CallMethod("IND_ExpenseCreation.justificationPanel",Ereason,Vname,GSTIN,I_no,I_Date);
}
 

}

}
if(!addedlines)
ValidationUtils.verify(false,true,"WorkCode is not availble in to Create Budget");
else{ 
 var document = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Documents");
 

WorkspaceUtils.waitForObj(document);
document.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var attchDocument = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl","New");
  WorkspaceUtils.waitForObj(attchDocument);
  ReportUtils.logStep_Screenshot();
  attchDocument.Click();
  aqUtils.Delay(4000, "Waiting to Open file");;
  var dicratory = eval(WorkspaceUtils.Sys_Maconomy_Parent).Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = eval(WorkspaceUtils.Sys_Maconomy_Parent).Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Document Attached");
  TextUtils.writeLog("Document Attached");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var submit = getObjectAddress_JavaClasssName(Maconomy_ParentAddress,"SingleToolItemControl","Submit");

Sys.HighlightObject(submit)
  WorkspaceUtils.waitForObj(submit);
  ReportUtils.logStep_Screenshot();
  submit.Click();
  waitUntil_MaconomyScreen_loaded_Completely();
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Expense Number",EnvParams.Opco,"Data Management",ExpenseNumber);
  ExcelUtils.WriteExcelSheet("Expense Description",EnvParams.Opco,"Data Management",desp);
  TextUtils.writeLog("Created Expenses Number :"+ExpenseNumber);

  }
  }

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrlc
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}








