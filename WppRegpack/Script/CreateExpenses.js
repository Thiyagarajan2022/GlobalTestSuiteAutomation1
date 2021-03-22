//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT EventHandler
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
var employeeNo,Description,VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice,NOL = "";
var Language = "";



var Arrays = [];
var count = true;
var STIME = "";
var Description;
var jobNumber = "";
var Language = "";

function CreateExpense() {
TextUtils.writeLog("Create Purchase Order Started"); 
Indicator.PushText("waiting for window to open");

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

Arrays = [];
count = true;
STIME = "";
Description;
jobNumber = "";

          STIME = WorkspaceUtils.StartTime();
          getDetails();
          goToJobMenuItem();
          CreateEmpxnse();
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



//Go To Job from Menu
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
TextUtils.writeLog("Entering into Time & Expenses from Time & Expenses Menu");
}



function CreateEmpxnse(){ 
  ReportUtils.logStep("INFO", "Enter Expenses Details");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var ExpenseTab = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  waitForObj(ExpenseTab)
  ReportUtils.logStep_Screenshot("");
  ExpenseTab.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
//  if(ImageRepository.ImageSet.Close_Filter.Exists()){
//  ImageRepository.ImageSet.Close_Filter.Click();
//  ReportUtils.logStep_Screenshot("");
//  }else{

//var closeFilter = "";
//if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.isVisible()){
//closeFilter = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
//Sys.HighlightObject(closeFilter)
//}
//else{ 
//closeFilter = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//Sys.HighlightObject(closeFilter)
//}
////  var closeFilter = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
//  Log.Message(closeFilter.FullName)
//  waitForObj(closeFilter)
//  ReportUtils.logStep_Screenshot("");
//  closeFilter.Click();
////  }
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }
                  
var expenses = "";
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.isVisible()){
expenses =  Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(expenses)
}
else{ 
expenses =  Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
Sys.HighlightObject(expenses)
}

//  var expenses =  Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Log.Message(expenses.FullName)
  waitForObj(expenses)
  WorkspaceUtils.waitForObj(expenses);
  ReportUtils.logStep_Screenshot("");
  expenses.Click();
  TextUtils.writeLog("Create New Expense Sheet is Clicked");
  
//for(var i=0;i<expenses.ChildCount;i++){
//  Log.Message(expenses.Child(i).Name)
//  Log.Message(expenses.Child(i).toolTipText.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Expense Sheet (Ctrl+N)").OleValue.toString().trim())!=-1)
//  
//if((expenses.Child(i).isVisible())&&(expenses.Child(i).toolTipText.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Expense Sheet (Ctrl+N)").OleValue.toString().trim())!=-1)){
//  expenses = expenses.Child(i);
//  WorkspaceUtils.waitForObj(expenses);
//  ReportUtils.logStep_Screenshot("");
//  expenses.Click();
//  break;
//}    
//} 
  aqUtils.Delay(5000, "Create Expenses Sheet");
  
  var Cancel = Aliases.Maconomy.ExpenseSheet.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
  waitForObj(Cancel)

  var employee = Aliases.Maconomy.ExpenseSheet.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
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
  
  var descrip = Aliases.Maconomy.ExpenseSheet.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  descrip.HoverMouse();
  Sys.HighlightObject(descrip)
  WorkspaceUtils.waitForObj(descrip)
  descrip.setText(Description+" "+STIME); 
  
  var job = Aliases.Maconomy.ExpenseSheet.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.McValuePickerWidget;
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
  
  var Create = Aliases.Maconomy.ExpenseSheet.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim())
  WorkspaceUtils.waitForObj(descrip);
  ReportUtils.logStep_Screenshot(""); 
  Create.Click();
  TextUtils.writeLog("Expense Sheet is Created");
}



function gotoTimeExpenses(){
var ExpenseNumber = "";
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.isVisible()){
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.isVisible())
 ExpenseNumber = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
//var ExpenseNumber = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite;
Log.Message(ExpenseNumber.FullName)
Sys.HighlightObject(ExpenseNumber)
}else{
 ExpenseNumber = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
Log.Message(ExpenseNumber.FullName)
Sys.HighlightObject(ExpenseNumber)
}
//var ExpenseNumber = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
WorkspaceUtils.waitForObj(ExpenseNumber);
ExpenseNumber = ExpenseNumber.getText().OleValue.toString().trim();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
//var addline = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//WorkspaceUtils.waitForObj(addline);
//addline.Click();

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
 
var addline = "";
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
 addline = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(addline)
}
else{
 addline = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(addline)
}
//var addline = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
WorkspaceUtils.waitForObj(addline);
ReportUtils.logStep_Screenshot();
addline.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  
var EntryDate = "";
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
  EntryDate = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
Sys.HighlightObject(EntryDate)
}
else{
 EntryDate = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;;
Sys.HighlightObject(EntryDate)
}
//var EntryDate = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
WorkspaceUtils.waitForObj(EntryDate);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    
var workCode = ""; 
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
  workCode = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(workCode)
}
else{
 workCode = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(workCode)
}  
//var workCode = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
workCode.Click();
WorkspaceUtils.SearchByValue(workCode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"Work Code :"+wCodeID);
Sys.HighlightObject(workCode);
var Wdes = workCode.getText();
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;

var WDesp = ""; 
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
  WDesp = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(WDesp)
}
else{
 WDesp = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(WDesp)
} 
//var WDesp = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
WDesp.setText(Wdes);
Sys.HighlightObject(WDesp);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    
var Currency = ""; 
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
  Currency = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
Sys.HighlightObject(Currency)
}
else{
 Currency = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
Sys.HighlightObject(Currency)
} 
//var Currency = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
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
    
var UnitPrice = ""; 
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
  UnitPrice = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(UnitPrice)
}
else{
 UnitPrice = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(UnitPrice)
} 

//var UnitPrice = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(UnitPrice);
UnitPrice.setText(Amt);

var save = ""; 
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
  save = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(save)
}
else{
 save = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(save)
} 
//var save = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
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
 var document = "";
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
 document = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(document)
}
else{
 document = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(document)
} 
//var document = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(document);
document.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var attchDocument = "";
if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.isVisible()){
 attchDocument = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(attchDocument)
}
else{
 attchDocument = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
Sys.HighlightObject(attchDocument)
}
//var attchDocument = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  WorkspaceUtils.waitForObj(attchDocument);
  ReportUtils.logStep_Screenshot();
  attchDocument.Click();
  aqUtils.Delay(4000, "Waiting to Open file");;
  var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Document Attached");
  TextUtils.writeLog("Document Attached");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var submit = "";

//submit = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Submit
//if(Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.isVisible()){
// submit = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//Sys.HighlightObject(submit);
//}
//else{
// submit = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
//Sys.HighlightObject(submit);
//}

//  var submit = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;


var table = "";
var Add = [];
var childcount = 0;
var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
Sys.Process("Maconomy").Refresh()
for(var i = 0;i<Parent.ChildCount;i++){ 
  if((Parent.Child(i).isVisible()) && (Parent.Child(i).ChildCount == 3)){
  Add[childcount] = Parent.Child(i);
  childcount++;
  }
}

Parent = "";
var pos = 0;
for(var i=0;i<Add.length;i++){ 
  if(Add[i].Height>pos){ 
    pos = Add[i].Height;
    Parent = Add[i];
  }
}


Log.Message(Parent.FullName)
submit = Parent.SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
Sys.HighlightObject(submit)
  WorkspaceUtils.waitForObj(submit);
  ReportUtils.logStep_Screenshot();
  submit.Click();
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Expense Number",EnvParams.Opco,"Data Management",ExpenseNumber);
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








