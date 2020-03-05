﻿//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT TestRunner
//USEUNIT Expenses


var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "CreateExpense";
  Indicator.Show();
  Indicator.PushText("waiing for window to open");

Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
Log.Message(sheetName);
var Arrays = [];
var count = true;
var STIME = "";
var Description;
var jobNumber = "";

//function mainscript(){ 
//Job_name =   ReadLog.readlog();
//aqUtils.Delay(1000, Indicator.Text);;;
//}

var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "CreateExpense", true);

function getDetails(){
ExcelUtils.setExcelName(workBook, "Server Details", true);
var employeeNo = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
Log.Message(employeeNo);

ExcelUtils.setExcelName(workBook, sheetName, true);
Description= ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
if((Description==null)||(Description=="")){ 
ValidationUtils.verify(false,true,"Description is Needed to Create a Expenses");
Log.Message(Description);
}
}


////------Label Validating Field-----////

function address(){
Delay(4000);
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


function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
    menuBar.DblClick();
     if(ImageRepository.ImageSet2.TE.Exists()){
       ImageRepository.ImageSet2.TE.Click();// GL
      }
     else if(ImageRepository.ImageSet2.TE1.Exists()){
       ImageRepository.ImageSet2.TE1.Click();
      }
     else{
       ImageRepository.ImageSet2.TE2.Click();
    }
    
      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
      var Modify_Budget;
       for(var i=1;i<=childCC;i++){ 
          Time_Expenses = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i);
          if(Time_Expenses.isVisible()){ 
          Time_Expenses = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", ""); 
          Time_Expenses.DblClickItem("|Time & Expenses");
       }
      }
     Delay(6000);
     
     ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
}

function gotoTimeExpenses(){
    ReportUtils.logStep("INFO", "Enter Expenses Details");
  Delay(2000);
  var expenses =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  Sys.HighlightObject(expenses);
  expenses.Click();
  Delay(4000);
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).Refresh();
   var grid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1);
  
   
    var linestatus = false;
  if(!linestatus)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").isVisible())
  {
  var newsheet =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  linestatus = true;
   } 
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").isVisible())
  {
  var newsheet =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  linestatus = true;
   }
//   if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").isVisible())
//  {
//  var newsheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
//  linestatus = true;
//   }
   Sys.HighlightObject(newsheet);
   newsheet.Click(); 

  Delay(1000);
  address();
  
//-----Entering Employee details-----////      
  
  ///----From Notepad------  
//  ExcelUtils.setExcelName(workBook, "Server Details", true);
//  var employeeNo = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//  Log.Message(employeeNo);
//      var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2)
//      Log.Message(employee.getText());
//      if(employee.getText()!=employeeNo){
//      if(employeeNo!=""){
//        employee.Click();
//        WorkspaceUtils.SearchByValueTable(employee,"Employee",employeeNo,"Employee Number");
//      }
//      }
//      else{
//        ValidationUtils.verify(true,true,"Employee Number is Exist in the Create Expenses");
//      }  
          
     ////----From Excel------ 
      ExcelUtils.setExcelName(workBook, sheetName, true);
    var employeeNo = ExcelUtils.getColumnDatas("Employeeno",EnvParams.Opco)
    Log.Message(employeeNo);
      var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2)
      Log.Message(employee.getText());
      if(employee.getText()!=employeeNo){
      if(employeeNo!=""){
        employee.Click();
        WorkspaceUtils.SearchByValueTable(employee,"Employee",employeeNo,"Employee Number");
      }
      }
      else{
        ValidationUtils.verify(true,true,"Employee Number is Exist in the Create Expenses");
      }         

    
////-----Entering Description ----//

var descrip = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  descrip.setText(Description+" "+STIME); 
  Delay(3000)  
//  if((descrip.getText().OleValue.toString().trim()==null)||(descrip.getText().OleValue.toString().trim()==""))
//  ValidationUtils.verify(false,true,"Description can't able to enter in Maconomy");
//  else
//  ValidationUtils.verify(true,true,"Description is enter in Maconomy");
  

////-----Entering Jobname ----//
  jobNumber = readlog();
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(jobNumber!=""){
   job.Click();
   Log.Message(jobNumber);
      WorkspaceUtils.SearchByValues(job,"Job",jobNumber,"Job Number");
  } 
  else{ 
  ValidationUtils.verify(false,true,"Job Number is Exist in the Create Expenses");
  } 

   
  var createbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create"); 
  Sys.HighlightObject(createbtn);
  if(createbtn.isEnabled()){    
    Log.Message("Create button is visible");
  Log.Message("Expenses is Created");
    createbtn.Click();
    ReportUtils.logStep("INFO",Description+" "+STIME +" : is Created");
  } 
  else{
    var cancelbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
    Sys.HighlightObject(cancelbtn)    
    cancelbtn.Click();
    ReportUtils.logStep("ERROR","Expenses is not Created");
  } 
  Delay(4000); 
  
//// ------ Get created Expenses sheet number ----//
  
 
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").Index==2)
  var get = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3).getText();
  else
  var get = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3).getText();
  Log.Message("Expense Sheet :" + get);
  ValidationUtils.verify(true,true,"Created Expenses Number : "+get);
  
////------------Verify the Expenses Available in the list    
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible())
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  else
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible())
  var allexpenses = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Expense Sheets");
  else
  var allexpenses = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Expense Sheets");
  allexpenses.Click();
  
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible())
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
  else
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
  firstcell.Click();
  firstcell.setText(get);
  firstcell.Keys("[Tab]");

  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible())
  var des = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
  else
  var des = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
  des.Click();
  des.setText(Description+" "+STIME);
  Delay(2000);
  var flag =false;
  for(var i=0;i<table.getItemCount();i++){
    if(table.getItem(i).getText_2(2).OleValue.toString().trim()==(Description+" "+STIME)){
      flag = true;
      Log.Message("Created Expenses Sheet is listed Below :" +Description);
//      var path = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+TestRunner.sheet.substring(TestRunner.sheet.lastIndexOf("\\")+1,TestRunner.sheet.indexOf("."))+".txt";
//      TextUtils.writeDetails(path,"Expenses Number ",get);
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
         Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
      break;
    } 
    else{
      table.Keys("[Down]");
    } 
  }   
  
  ValidationUtils.verify(flag,true,"Created Expenses is available in system");
  ValidationUtils.verify(true,true,"Expenses Number :"+table.getItem(i).getText_2(2).OleValue.toString().trim())
  ReportUtils.logStep("INFO", "Created Expenses is listed in table"); 
  }    
  
  
///-----Registration tab-------/////

  function gotoregister(){  
  jobNumber = readlog();
  ReportUtils.logStep("INFO", "Entering Expenses details in Registration Tab");
    
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    ref.Refresh();
    var addedlines = false;   
//    for(var z=1;z<=10;z++){
//    sheetName ="CreateExpense";
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    var workcode = ExcelUtils.getColumnDatas("Workcode_"+z,EnvParams.Opco)
//    var DetailDescription = ExcelUtils.getColumnDatas("Detail Description_"+z,EnvParams.Opco)
//    var Description = ExcelUtils.getColumnDatas("Description_"+z,EnvParams.Opco)
//    var jobname = ExcelUtils.getColumnDatas("JobName_"+z,EnvParams.Opco)
//    var currency = ExcelUtils.getColumnDatas("currency_"+z,EnvParams.Opco)
//    var Amount =  ExcelUtils.getColumnDatas("Amount_"+z,EnvParams.Opco)
      
//     var linestatus = false;    
//     ref.Refresh();
//     
//      
//    var linestatus = false;
//    if(!linestatus) 
//    if((workcode!="")&&(workcode!=null)){  
//    if(workcode!=""){    
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
//    {
//    var Addbutton = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
//    linestatus = true;
//    }
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
//    {
//     var Addbutton = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
//    linestatus = true; 
//    }
//    Delay(3000);
//    Addbutton.Click();
//    Delay(5000);
//    
//    linestatus = false; 
// 
//   var commAdd ; 
//if(!linestatus){
//      if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible()){
//    var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).ChildCount;
//    for(var i =0;i<jobno;i++){ 
//    var job1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3);
//    Log.Message(job1.Child(i).Name);
//    if((job1.Child(i).Name.indexOf("McClumpSashForm")!=-1) && (job1.Child(i).isVisible())){
//       
//    if(job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).Name.indexOf("McTableWidget")!=-1){
//    var commAdd = job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).SWTObject("McGrid", "", 2);
//    linestatus = true;
//    } }}}} 
//if(!linestatus){
//      if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible()){
//    var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).ChildCount;
//    for(var i =0;i<jobno;i++){ 
//    var job1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3);
//    Log.Message(job1.Child(i).Name);
//    if((job1.Child(i).Name.indexOf("McClumpSashForm")!=-1) && (job1.Child(i).isVisible())){
//       
//    if(job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).Name.indexOf("McTableWidget")!=-1){
//    var commAdd = job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).SWTObject("McGrid", "", 2);
//linestatus = true;
//} }  } }  } 
//if(!linestatus){
//      if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible()){
//    var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).ChildCount;
//    for(var i =0;i<jobno;i++){ 
//    var job1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3);
//    Log.Message(job1.Child(i).Name);
//    if((job1.Child(i).Name.indexOf("McClumpSashForm")!=-1) && (job1.Child(i).isVisible())){
//       
//    if(job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).Name.indexOf("McTableWidget")!=-1){
//    var commAdd = job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).SWTObject("McGrid", "", 2);
//linestatus = true;
//} } } } }    
//    
//    
// var Ref =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
//    
//    
//    
/////--- Adding lines in registration tab----
//    
////    if(!linestatus){
//    
//    var TodayValue = aqDateTime.Today();
//    var StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
//    Log.Message(StringTodayValue);
//    Delay(1000);
//    Sys.Desktop.KeyDown(0x09); // Press Ctrl
//    Sys.Desktop.KeyUp(0x09);
//    
//////---Checking Job Number      
//
//    
//Ref.Refresh();  
//    var job ;
//Log.Message(commAdd.FullName);
//Sys.HighlightObject(commAdd)
//    var job_name = commAdd.ChildCount;
//    aqUtils.Delay(1000, Indicator.Text);;;
//    for(var jId=0;jId<job_name;jId++){
//    if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McValuePickerWidget")!=-1)){
//     job =  commAdd.Child(jId);
//     Log.Message(job.FullName);
//     }}
//    Log.Message(job.getText());  
//    if(job.getText()!=jobNumber){     
//      if(jobNumber!=""){
//      job.Click();    
//      WorkspaceUtils.SearchByValuesjob(job,"Job",jobNumber,"Job Number")    
//      }
//    }
//     else{
//          ValidationUtils.verify(true,true,"Job Number is Exist in the Create a Expense Sheet");  
//        }
//    Delay(1000);
//    job.Keys("[Tab]");
//    Delay(1000);   
//   
//        
/////----Entering Workcode----   
//Ref.Refresh();    
//    var workCode ;
//    var job_name = commAdd.ChildCount;
//    aqUtils.Delay(1000, Indicator.Text);;
//    for(var jId=0;jId<job_name;jId++){
//    aqUtils.Delay(1000, Indicator.Text);;
//    if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McValuePickerWidget")!=-1)){
//     workCode =  commAdd.Child(jId);
//     Log.Message(workCode.FullName);
////     linestatus = true;
//     }}
//    
//    if(workcode!=""){
//   addedlines = true;
//    workCode.Click();
//    WorkspaceUtils.SearchByValue(workCode,"Work Code",workcode,"Work Code :"+workcode);
//           }else{ 
//      ValidationUtils.verify(false,true,"WorkCode Needed to Create Expenses");
//    }  
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    Delay(1000);   
//           
//    
//Ref.Refresh();    
//    var detaildescription ;
//    var job_name = commAdd.ChildCount;
//    Log.Message(job_name)
//    aqUtils.Delay(1000, Indicator.Text);;
//    for(var jId=0;jId<job_name;jId++){
//    Log.Message(commAdd.Child(jId).Name)
//    aqUtils.Delay(1000, Indicator.Text);;
//    if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McTextWidget")!=-1)){
//     detaildescription =  commAdd.Child(jId);
//     Log.Message(detaildescription.FullName);
////     linestatus = true;
//     }}
//    
//    
//    Sys.HighlightObject(detaildescription);
//    detaildescription.Click();
//    if(DetailDescription!=""){
//    detaildescription.setText(DetailDescription);
//    ValidationUtils.verify(true,true,"Detail Description is Entered");
//    }
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    Delay(2000);
//////----Selecting Currency from DropDown----   
// 
//Ref.Refresh();    
//    var currency1 ;
//    var job_name = commAdd.ChildCount;
//    aqUtils.Delay(1000, Indicator.Text);;
//    for(var jId=0;jId<job_name;jId++){
//    aqUtils.Delay(1000, Indicator.Text);;
//    if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McPopupPickerWidget")!=-1)){
//     currency1 =  commAdd.Child(jId);
//     Log.Message(currency1.FullName);
////     linestatus = true;
//     }}
//    
//    
//    currency1.Keys(" ");    
//    if(currency!=""){
//       currency1.Click();
//       WorkspaceUtils.DropDownList(currency,currency1);
//       Delay(1000); 
//    } 
//    else{
//      ValidationUtils.verify(false,true,"Currency is Needed to Create a Expense Sheet");  
//    } 
//    Delay(1000);
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);      
//    
///////-------Entering Amount      
//Ref.Refresh();      
//    var amounnt ;
//    var job_name = commAdd.ChildCount;
//    aqUtils.Delay(1000, Indicator.Text);;
//    for(var jId=0;jId<job_name;jId++){
//    aqUtils.Delay(1000, Indicator.Text);;
//    if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McTextWidget")!=-1)){
//     amounnt =  commAdd.Child(jId);
//     Log.Message(amounnt.FullName);
////     linestatus = true;
//     }}
//    
//    
//    Sys.HighlightObject(amounnt);
//    amounnt.Click();
//    if(Amount!=""){
//      amounnt.setText(Amount);
//      ValidationUtils.verify(true,true,"Amount is Entered");
//    } 
//    Delay(2000);
//
/////--- Entering Save button-----
//
//    var linestatus = false;
//    if(!linestatus) 
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
//    {
//    var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//    linestatus = true;
//    }
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
//    {
//    var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//    linestatus = true; 
//    }
//    save.Click();
//    Delay(5000);
//     
//    ValidationUtils.verify(true,true,"Expense line is added and saved");
//    Delay(5000);   
//    }
//       
//    else{
//      ValidationUtils.verify(false,true,"Data Needed to addline Expenses");
//    }    
//    }
//    }
//
//    
//
/////---Clicking Dowcumnet Tab   
    var linestatus = false;
    if(!linestatus) 
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
    {
    var documents = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    linestatus = true;
    }
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
    {
    var documents = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    linestatus = true; 
    }
    documents.Click();
    Delay(5000);  
////-------New icon button    
   var linestatus = false;
    
   if(!linestatus)  
   if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
   {
   var newicon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2); 
   linestatus = true; 
    }      
    if(!linestatus)  
   if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible())
   {   
   var newicon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
   linestatus = true; 
    }
    if(!linestatus)  
   if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
   {   
   var newicon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
   linestatus = true; 
    }
  newicon.Click();
  Delay(2000);
  
   var uploadlocal =  Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1); 
          uploadlocal.Keys("C:\\Users\\674087\\Desktop\\test.txt");  
//  var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";
//      Log.Message("Notepad :"+notepadPath)
//          uploadlocal.Keys(TestRunner.sheet);
          Sys.Desktop.KeyDown(0x0D);
          Sys.Desktop.KeyUp(0x0D); 
            Delay(8000);
            ValidationUtils.verify(true,true,"Document is Uploaded")
       
    var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
    Sys.HighlightObject(submit);
    submit.Click();  
//    ValidationUtils.verify(true,true,"Expense Sheet is Submitted")
//    Log.Message("Expense Sheet is Submitted"); 
//    Delay(3000);
//    var print = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 5000);
//    print.Click();
//    Delay(5000);   
    
//    var printstatus = false;
//    if(!printstatus)
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).isVisible())
//    {
//    var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 7000);
//    printstatus = true;
//    }
//    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible())
//    {
//    var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 7000);
//    printstatus = true;
//    } 

   
    Sys.HighlightObject(submit);
    submit.Click();  
    ValidationUtils.verify(true,true,"Expense Sheet is Submitted")
    Log.Message("Expense Sheet is Submitted"); 
    Delay(3000);
    var printstatus = false;
    if(!printstatus)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible())
    {
    var print = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 5000);
    printstatus = true;
    }
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).isVisible())
    {
    var print = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 5000);
    printstatus = true;
    }  
    Sys.HighlightObject(print);
    print.Click();
    Delay(5000);
  }






function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrlc
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}




function excel(start){ 

var Arrayss = [];
var xlDriver = DDT.ExcelDriver(TestRunner.sheet, sheetName, true);
var id =0;
var colsList = [];
Log.Message(DDT.CurrentDriver.ColumnCount);
   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
   xlDriver.Next();

     while (!DDT.CurrentDriver.EOF()) {
    var temp ="";
      for(var idx=start;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
     if(temp.length!=10){
     Arrayss[id]=temp;
     Log.Message(temp)
     }
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return Arrayss;
}


//var excelName ="C:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\Regression\\China\\TESTAPAC\\TESTAPAC_China.xlsx";
//sheet = "CreateExpense";


function CreateExpense() {
      Language = EnvParams.Language;
        if((Language==null)||(Language=="")){
          ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
        }
      Log.Message(Language)

          STIME = WorkspaceUtils.StartTime();
          ReportUtils.logStep("INFO", "Expenses Creation started::"+STIME);
          getDetails();
//          goToJobMenuItem();
//          gotoTimeExpenses();
          gotoregister();
          WorkspaceUtils.closeAllWorkspaces();

}







function SOXexcel(start){ 
var Arrayss = []; 
//var xlDriver = DDT.ExcelDriver(TestRunner.sheet, sheetName, true);
var xlDriver = DDT.ExcelDriver(workbok, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
     Arrayss[id]=temp;
     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrayss;
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

function readlog(){ 

sheetName = "JobCreation";
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
Log.Message("Notepad :"+notepadPath)
return TextUtils.readDetails(notepadPath,"Job Number");
//Log.Message( readDetails("C:\\Users\\674087\\Documents\\TestComplete 14 Projects\\After Stuart Discussion\\WppRegression_v12.50\\WppRegPack\\RegressionLogs\\TESTAPAC\\Regression\\China\\1307-Sudler China(MDS)_Client Billable.txt","Job Number") );
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
var id =0;
var colsList = [];
var temp ="";
     while (!DDT.CurrentDriver.EOF()) {
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf("*")!=-1){
     var excelData =  temp.split("*");
      
     }else if(temp.length>0){ 
      excelData[0] = temp;
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

















