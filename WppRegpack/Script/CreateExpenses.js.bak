//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT TestRunner
//USEUNIT Expenses
//USEUNIT Restart

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "CreateExpense";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");

ExcelUtils.setExcelName(workBook, sheetName, true);

var Arrays = [];
var count = true;
var STIME = "";
var Description;
var jobNumber = "";
var Language = "";


var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "CreateExpense", true);

function getDetails(){
    ExcelUtils.setExcelName(workBook, "Server Details", true);
    var employeeNo = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

    ExcelUtils.setExcelName(workBook, sheetName, true);
    Description= ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
    if((Description==null)||(Description=="")){ 
    ValidationUtils.verify(false,true,"Description is Needed to Create a Expenses");
    Log.Message(Description);
    }
    
    jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
    if((jobNumber=="")||(jobNumber==null)){
    ExcelUtils.setExcelName(workBook, "Data Management", true);
    jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
    }  
    if((jobNumber=="")||(jobNumber==null))
    ValidationUtils.verify(false,true,"Job Number is needed to Create Budget");
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


function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
    menuBar.DblClick();
     if(ImageRepository.ImageSet0.TimeExpense.Exists()){
       ImageRepository.ImageSet0.TimeExpense.Click();// GL
      }
     else if(ImageRepository.ImageSet0.TimeExpense1.Exists()){
       ImageRepository.ImageSet0.TimeExpense1.Click();
      }
     else{
       ImageRepository.ImageSet0.TimeExpense2.Click();
    }

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
Delay(3000);
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
Client_Managt.ClickItem("|Time & Expenses");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Time & Expenses");
}

}    

     aqUtils.Delay(5000, Indicator.Text);     
     ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
}

function gotoTimeExpenses(){
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    ReportUtils.logStep("INFO", "Enter Expenses Details");
  var expenses =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  waitForObj(expenses)
  Sys.HighlightObject(expenses);
  expenses.HoverMouse();
  expenses.Click();
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).Refresh();
   var grid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1);
    grid.HoverMouse();
    Sys.HighlightObject(grid);
    var linestatus = false;
  if(!linestatus)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").isVisible())
  {
  var newsheet =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  linestatus = true;
   } 
   if(!linestatus)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").isVisible())
  {
  var newsheet =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  linestatus = true;
   }
   if(!linestatus)
   if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").isVisible())
  {
  var newsheet = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  linestatus = true;
   }
   Sys.HighlightObject(newsheet);
       newsheet.HoverMouse();
  ReportUtils.logStep_Screenshot("");
   newsheet.Click(); 
  address();
  
//-----Entering Employee details-----////      
  
     ////----From Excel------ 
      ExcelUtils.setExcelName(workBook, sheetName, true);
    var employeeNo = ExcelUtils.getColumnDatas("Employeeno",EnvParams.Opco)
     var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2)
      if(employee.getText()!=employeeNo){
      if(employeeNo!=""){
        employee.HoverMouse();
        Sys.HighlightObject(employee);
         employee.HoverMouse();
        employee.Click();
        WorkspaceUtils.SearchByValueTable(employee,"Employee",employeeNo,"Employee Number");
      }
      }
      else{
        ValidationUtils.verify(true,true,"Employee Number is Exist in the Create Expenses");
      } 
          

    
////-----Entering Description ----//

  var descrip = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  descrip.HoverMouse();
  Sys.HighlightObject(descrip)
  descrip.HoverMouse();
  descrip.setText(Description+" "+STIME); 
  
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  job.HoverMouse();
  Sys.HighlightObject(job)
  job.HoverMouse();
  if(job.getText()!=jobNumber){
  if(jobNumber!=""){
   job.Click();
      WorkspaceUtils.SearchByValues(job,"Job",jobNumber,"Job Number");
  } 
  }
  else{ 
  ValidationUtils.verify(false,true,"Job Number is Exist in the Create Expenses");
  } 

 
  var createbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create"); 
  Sys.HighlightObject(createbtn);
  if(createbtn.isEnabled()){   
    createbtn.HoverMouse();
    ReportUtils.logStep_Screenshot(""); 
    createbtn.Click();
    TextUtils.writeLog("Expenses Sheet is CREATED");
    ValidationUtils.verify(true,true,"Expenses is CREATED");
    ReportUtils.logStep("INFO",Description+" "+STIME +" : is Created");
  } 
  else{
    var cancelbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Expense Sheet").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
    Sys.HighlightObject(cancelbtn)    
    cancelbtn.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    cancelbtn.Click();
    TextUtils.writeLog("Expenses Sheet is not CREATED");
    ValidationUtils.verify(true,false,"Expenses is not Created");
    ReportUtils.logStep("ERROR","Expenses is not Created");
  } 
  aqUtils.Delay(5000, "Getting Expenses Number");
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").Index==2)
  var get = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3).getText();
  else
  var get = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3).getText();
  Log.Message("Expense Sheet :" + get);
  ValidationUtils.verify(true,true,"Created Expenses Number : "+get);
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Expense Number",EnvParams.Opco,"Data Management",get)
  
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
  allexpenses.HoverMouse();
  ReportUtils.logStep_Screenshot("");
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
  aqUtils.Delay(2000, Indicator.Text);
  var flag =false;
  for(var i=0;i<table.getItemCount();i++){
    if(table.getItem(i).getText_2(2).OleValue.toString().trim()==(Description+" "+STIME)){
      flag = true;
      ReportUtils.logStep_Screenshot("");
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
  TextUtils.writeLog("Created Expenses is available in system");
  ValidationUtils.verify(flag,true,"Created Expenses is available in system");
  ValidationUtils.verify(true,true,"Expenses Number :"+table.getItem(i).getText_2(2).OleValue.toString().trim())
  ReportUtils.logStep("INFO", "Created Expenses is listed in table"); 
  }    
  
  
///-----Registration tab-------/////

  function gotoregister(){  
    
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    ref.Refresh();
    
    var addedlines = false;   
    for(var z=1;z<=10;z++){
    sheetName ="CreateExpense";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    
    var workcode = ExcelUtils.getColumnDatas("Work Description_"+z,EnvParams.Opco)
    var currency = ExcelUtils.getColumnDatas("currency_"+z,EnvParams.Opco)
    var Amount =  ExcelUtils.getColumnDatas("Amount_"+z,EnvParams.Opco)
    var Reason = ExcelUtils.getColumnDatas("Expense Reason_"+z,EnvParams.Opco)
    var Gstin = ExcelUtils.getColumnDatas("GSTIN_"+z,EnvParams.Opco)
    var InvoiceNo =  ExcelUtils.getColumnDatas("Invoice No_"+z,EnvParams.Opco)
    var InvoiceDate =  ExcelUtils.getColumnDatas("Invoice Date_"+z,EnvParams.Opco)
    var VendorName =  ExcelUtils.getColumnDatas("Vendor Name_"+z,EnvParams.Opco)
           
    ref.Refresh();
    var linestatus = false;
    
    if((workcode!="")&&(workcode!=null)){  
        if(workcode!=""){    
        if(!linestatus) 
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
        {
        var Addbutton = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
        linestatus = true;
        }
        if(!linestatus) 
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
        {
         var Addbutton = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
        linestatus = true; 
        }
        Sys.HighlightObject(Addbutton);
        Addbutton.HoverMouse();
        ReportUtils.logStep_Screenshot("");
        Addbutton.Click();    
        linestatus = false; 
        aqUtils.Delay(2000, Indicator.Text);
       var commAdd ; 
      if(!linestatus){
          if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible()){
       var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).ChildCount;
        for(var i =0;i<jobno;i++){ 
        var job1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3);
    //    Log.Message(job1.Child(i).Name);
        if((job1.Child(i).Name.indexOf("McClumpSashForm")!=-1) && (job1.Child(i).isVisible())){
       
        if(job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).Name.indexOf("McTableWidget")!=-1){
        var commAdd = job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).SWTObject("McGrid", "", 2);
        linestatus = true;
        } }}}} 
      if(!linestatus){
          if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible()){
        var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).ChildCount;
        for(var i =0;i<jobno;i++){ 
        var job1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3);
        if((job1.Child(i).Name.indexOf("McClumpSashForm")!=-1) && (job1.Child(i).isVisible())){
       
        if(job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).Name.indexOf("McTableWidget")!=-1){
        var commAdd = job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).SWTObject("McGrid", "", 2);
    linestatus = true;
    } }  } }  } 
    if(!linestatus){
          if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible()){
        var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).ChildCount;
        for(var i =0;i<jobno;i++){ 
        var job1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3);
        if((job1.Child(i).Name.indexOf("McClumpSashForm")!=-1) && (job1.Child(i).isVisible())){
       
        if(job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).Name.indexOf("McTableWidget")!=-1){
        var commAdd = job1.Child(i).SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").Child(0).SWTObject("McGrid", "", 2);
    linestatus = true;
    } } } } }    
    
    
     var Ref =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
    
    
    ///--- Adding lines in registration tab----
        
        var TodayValue = aqDateTime.Today();
        var StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
    //    Log.Message(StringTodayValue);
        aqUtils.Delay(1000, Indicator.Text);
        Sys.Desktop.KeyDown(0x09); // Press Ctrl
        Sys.Desktop.KeyUp(0x09);
  
        Ref.Refresh();  
        var job ;
        Sys.HighlightObject(commAdd)
        var job_name = commAdd.ChildCount;
        aqUtils.Delay(1000, Indicator.Text);;;
        for(var jId=0;jId<job_name;jId++){
        if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McValuePickerWidget")!=-1)){
         job =  commAdd.Child(jId);
         }}    
    
        if(job.getText()!=jobNumber){     
          if(jobNumber!=""){
          job.Click();    
          WorkspaceUtils.SearchByValuesjob(job,"Job",jobNumber,"Job Number")    
          }
        }
         else{
              ValidationUtils.verify(true,true,"Job Number is Exist in the Create a Expense Sheet");  
            }    
        job.Keys("[Tab]");
        aqUtils.Delay(1000, Indicator.Text);   
   
        
    ///----Entering Workcode----    
        

        Ref.Refresh();    
        var workCode ;
        var job_name = commAdd.ChildCount;
        for(var jId=0;jId<job_name;jId++){    
        if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McValuePickerWidget")!=-1)){
         workCode =  commAdd.Child(jId);
         }}
    
        if(workcode!=""){
        addedlines = true;
        workCode.Click();
        WorkspaceUtils.SearchByValue(workCode,"Work Code",workcode,"Work Code :"+workcode);
               }else{ 
          ValidationUtils.verify(false,true,"WorkCode Needed to Create Expenses");
        }  
        Sys.Desktop.KeyDown(0x09);
        Sys.Desktop.KeyUp(0x09);
        aqUtils.Delay(1000, Indicator.Text);

    ///// -----Entering Detail Description
           
    
    //    Ref.Refresh();    
    //    var detaildescription ;
    //    var job_name = commAdd.ChildCount;
    //    for(var jId=0;jId<job_name;jId++){
    //    aqUtils.Delay(1000, Indicator.Text);
    //    if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McTextWidget")!=-1)){
    //     detaildescription =  commAdd.Child(jId);
    //     }}
    //       
    //    Sys.HighlightObject(detaildescription);
    //    detaildescription.Click();
    //    if(DetailDescription!=""){
    //    detaildescription.setText(DetailDescription);
    //    ValidationUtils.verify(true,true,"Detail Description is Entered");
    //    }
        Sys.Desktop.KeyDown(0x09);
        Sys.Desktop.KeyUp(0x09);
        aqUtils.Delay(1000, Indicator.Text);
    
    ////----Selecting Currency from DropDown----         
        Ref.Refresh();    
        var currency1 ;
        var job_name = commAdd.ChildCount;
        for(var jId=0;jId<job_name;jId++){
        if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McPopupPickerWidget")!=-1)){
         currency1 =  commAdd.Child(jId);
         }}
     
    //     Sys.Process("Maconomy").Refresh();  
        currency1.Keys(" ");    
        if(currency!=""){
           currency1.Click();
           aqUtils.Delay(1000, Indicator.Text);
           WorkspaceUtils.DropDownList(currency,"Currency");
           aqUtils.Delay(1000, Indicator.Text); 
        } 
        else{
          ValidationUtils.verify(false,true,"Currency is Needed to Create a Expense Sheet");  
        } 
        aqUtils.Delay(1000, Indicator.Text);
        Sys.Desktop.KeyDown(0x09);
        Sys.Desktop.KeyUp(0x09);      
    
    /////-------Entering Amount    
     
      Ref.Refresh();      
        var amounnt ;
        var job_name = commAdd.ChildCount;
        for(var jId=0;jId<job_name;jId++){
        if((commAdd.Child(jId).isVisible())&&(commAdd.Child(jId).Name.indexOf("McTextWidget")!=-1)){
         amounnt =  commAdd.Child(jId);
         }}
    
    
        Sys.HighlightObject(amounnt);
        amounnt.Click();
        if(Amount!=""){
          amounnt.setText(Amount);
          ValidationUtils.verify(true,true,"Amount is Entered");
        } 
        aqUtils.Delay(2000, Indicator.Text);

    ///--- Entering Save button-----

        var linestatus = false;
        if(!linestatus) 
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
        {
        var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
        linestatus = true;
        }
        if(!linestatus) 
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
        {
        var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
        linestatus = true; 
        }
        if(!linestatus) 
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible())
        {
        var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
        linestatus = true; 
        }
        save.HoverMouse();
        ReportUtils.logStep_Screenshot("");
        save.Click();   
        TextUtils.writeLog("Expense line is added and saved");  
        ValidationUtils.verify(true,true,"Expense line is added and saved");
        aqUtils.Delay(3000, Indicator.Text); 
        }
       
      else{
        ValidationUtils.verify(false,true,"Data Needed to addline Expenses");
      }        
         if(EnvParams.Country.toUpperCase()=="INDIA"){
          Runner.CallMethod("IND_CreateExpenses.IndiaSpecific",Reason,Gstin,InvoiceDate,InvoiceNo,VendorName);
        }
    }    
   }
  
///---Clicking Dowcumnet Tab   
    var linestatus = false;
    if(!linestatus) 
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
    {
    var documents = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    linestatus = true;
    }
    if(!linestatus) 
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible())
    {
    var documents = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    linestatus = true; 
    }
    documents.HoverMouse();
    Sys.HighlightObject(documents);
    ReportUtils.logStep_Screenshot("");
    documents.Click();   
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
    newicon.HoverMouse();
    ReportUtils.logStep_Screenshot("");
      newicon.Click();
      aqUtils.Delay(1000, Indicator.Text);
   var uploadlocal =  Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1); 
      uploadlocal.Keys(workBook);
          Sys.Desktop.KeyDown(0x0D);
          Sys.Desktop.KeyUp(0x0D); 
            aqUtils.Delay(3000, "Attaching Document");
            TextUtils.writeLog("Document is Uploaded");
            ValidationUtils.verify(true,true,"Document is Uploaded")
      
    var submit = ""
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).ChildCount>=7){    
    submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
    }
    else if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).ChildCount>=7){    
    submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
    }   
    else if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).ChildCount>=7){    
    submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
    }
    
    Sys.HighlightObject(submit);
    submit.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    submit.Click();  
    TextUtils.writeLog("Expenses Sheet is Submitted");
    ValidationUtils.verify(true,true,"Expense Sheet is Submitted")
    aqUtils.Delay(2000, Indicator.Text);
  }

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrlc
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}


function CreateExpense() {
TextUtils.writeLog("Created Expenses is Started");
Indicator.PushText("waiing for window to open");
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);  
}
Arrays = [];
count = true;
STIME = "";
Description;
jobNumber = "";
Language = "";

          Language = EnvParams.Language;
          if((Language==null)||(Language=="")){
          ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
          }
          Language = EnvParams.LanChange(Language);
          WorkspaceUtils.Language = Language;
          STIME = WorkspaceUtils.StartTime();
          getDetails();
          goToJobMenuItem();
          gotoTimeExpenses();
          gotoregister();   
          WorkspaceUtils.closeAllWorkspaces();
}






