﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart
Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName ="";
var STIME = "";
var comapany,Job_group,quteNumber,ContryCurrency,ExchangeRate,ClientCurrency = "";
var jobNumber = "";
var Approve_Level = [];
var y=0;
var ApproveInfo = [];
var level =0;
var Language = "";

var Company_ID = "";
var Job_Name;
var templateJob="";
var WorkCode;
var Internal_Description;
var Line_Type;
var Employee_Categories;
var Employee_Number;
var Qly;
var CostBase;
var count = true;
var Arrays = [];
var workCodeList = [];
var workActivity = [];
var quteNumber ="";
var Jobnum = "";

function createBudget(sheet,Jobnumber){ 
  
TextUtils.writeLog("Job Budget Creation Started"); 
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

Log.Message("Language :"+Language)
excelName = EnvParams.path;
workBook = Project.Path+excelName;
STIME = "";
comapany,Job_group,ContryCurrency,quteNumber,ExchangeRate,ClientCurrency = "";
jobNumber = "";
templateJob="";
Approve_Level = [];
y=0;
ApproveInfo = [];
workCodeList = [];
workActivity = [];
level =0;
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
Jobnum = Jobnumber;
sheetName =sheet;
//try{
//sheetName = "JobCreation";
getDetails();
goToJobMenuItem();
goToBudget();
sheetName =sheet;
addingBudgetLines();
closeAllWorkspaces();
  
for(var i=level;i<ApproveInfo.length;i++){
level = i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
var temp = gg.split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
aprvBudget(temp[0],temp[1],temp[2]);
}

//}
//  catch(err){
//    Log.Message(err);
//  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
closeAllWorkspaces();
}




function getDetails(){ 

comapany = EnvParams.Opco
//sheetName ="JobBudgetCreation";
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  Log.Message("Job Number_"+Jobnum)
  jobNumber = ReadExcelSheet("Job Number_"+Jobnum,EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
//  sheetName ="JobBudgetCreation";
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Budget");

var CodeStatus = true;
var Country = EnvParams.Country;

 for(var i=1;i<=10;i++){
   Log.Message(sheetName)
ExcelUtils.setExcelName(workBook, sheetName, true);
var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
var Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
var Employee_Categories = ExcelUtils.getColumnDatas("Employee Categories_"+i,EnvParams.Opco)
var Employee_Number =  ExcelUtils.getColumnDatas("Employee Number_"+i,EnvParams.Opco)
var CostBase = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
var OutwardHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
var InwardHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)

if((wCodeID!="")&&(wCodeID!=null)){
  CodeStatus = false;
  if((Desp=="")||(Desp==null))
  ValidationUtils.verify(false,true,"Description_"+i+" is needed to Create Budget");

  if((Qly=="")||(Qly==null))
  ValidationUtils.verify(false,true,"Quantity_"+i+" is needed to Create Budget");
  
  if((CostBase=="")||(CostBase==null))
  ValidationUtils.verify(false,true,"Cost_"+i+" is needed to Create Budget");
 
  if(Country.toUpperCase()=="INDIA"){ 
  if((OutwardHSN=="")||(OutwardHSN==null))
  ValidationUtils.verify(false,true,"Outward HSN_"+i+" is needed to Create Budget");
  
  if((InwardHSN=="")||(InwardHSN==null))
  ValidationUtils.verify(false,true,"Inward HSN_"+i+" is needed to Create Budget");
  }
  
}
}

if(CodeStatus)
ValidationUtils.verify(false,true,"WorkCode is needed to Create Budget");

}

function goToJobMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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


function goToBudget(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var allJobs = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();

  var table = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
  
  var job = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
//  aqUtils.Delay(7000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to create budget");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy to create budget"); 
  closeFilter.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
var workCodeAdd = Aliases.Maconomy.WorkCodeValidation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(workCodeAdd);
workCodeAdd.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
workCodeList = [];
ExcelUtils.setExcelName(workBook, sheetName, true);
for(var i=1;i<=10;i++){

var temp = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco);
if(temp!=""){
workCodeList[i] = temp;
Log.Message(workCodeList[i])
}
}

workActivity = [];
var i=0
var WorkCodeGrid = Aliases.Maconomy.WorkCodeValidation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
for(var v=0;v<WorkCodeGrid.getItemCount();v++){ 
  for(var y=0;y<workCodeList.length;y++){ 
  if(WorkCodeGrid.getItem(v).getText(0).OleValue.toString().trim()==workCodeList[y]){ 
    workActivity[i] = WorkCodeGrid.getItem(v).getText(0).OleValue.toString().trim()+"*"+WorkCodeGrid.getItem(v).getText(6).OleValue.toString().trim()
    Log.Message(workActivity[i]);
    i++;
  }
  
  }
}



var Budget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
Log.Message(Budget.FullName)
WorkspaceUtils.waitForObj(Budget);
Budget.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
var show_budget = "";


//    show_budget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  var BarStat = true;
var show_budget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3;
for(var a=0;a<show_budget.ChildCount;a++){ 
  if((show_budget.Child(a).isVisible())&&(show_budget.Child(a).Name.indexOf("Composite")!=-1)&&(show_budget.Child(a).Index==1)){ 
    show_budget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
    Log.Message(show_budget.FullName);
    BarStat = false;
    break;
  }
}
if(BarStat){
var show_budget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5;
for(var a=0;a<show_budget.ChildCount;a++){ 
  if((show_budget.Child(a).isVisible())&&(show_budget.Child(a).Name.indexOf("Composite")!=-1)&&(show_budget.Child(a).Index==1)){ 
    show_budget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
    Log.Message(show_budget.FullName);
    break;
  }
}
}
    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);
    WorkspaceUtils.waitForObj(show_budget);
    show_budget.HoverMouse();
    show_budget.HoverMouse();
    show_budget.HoverMouse();

//    ClientCurrency = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  var BarStat = true;
ClientCurrency = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3;
for(var a=0;a<ClientCurrency.ChildCount;a++){ 
  if((ClientCurrency.Child(a).isVisible())&&(ClientCurrency.Child(a).Name.indexOf("Composite")!=-1)&&(ClientCurrency.Child(a).Index==1)){ 
    ClientCurrency = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
    Log.Message(ClientCurrency.FullName);
    BarStat = false;
    break;
  }
}
if(BarStat){
ClientCurrency = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5;
for(var a=0;a<ClientCurrency.ChildCount;a++){ 
  if((ClientCurrency.Child(a).isVisible())&&(ClientCurrency.Child(a).Name.indexOf("Composite")!=-1)&&(ClientCurrency.Child(a).Index==1)){ 
    ClientCurrency = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
    Log.Message(ClientCurrency.FullName);
    break;
  }
}
}
    Log.Message(ClientCurrency.FullName)
    WorkspaceUtils.waitForObj(ClientCurrency);
    ClientCurrency = ClientCurrency.getText();
    
    show_budget.Keys("Working Estimate"); 
    WorkspaceUtils.waitForObj(show_budget);
    aqUtils.Delay(2000,"Working Estimate is Selected");
    ValidationUtils.verify(true,true,"Working Estimate is Selected");
    TextUtils.writeLog("Working Estimate is Selected"); 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }else{ 
    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
    }

    }

    
}

function addingBudgetLines(){ 
  if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.isVisible()){
  var FullBudget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;
  }else{ 
  var FullBudget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl; 
  }
  Log.Message(FullBudget.FullName)
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  WorkspaceUtils.waitForObj(FullBudget);
  FullBudget.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var RowCount = 0;
var TotalBudget = 0.00;
var addedlines = false; 
  ExcelUtils.setExcelName(workBook, "CountryCurrency", true);
  ContryCurrency = ExcelUtils.getRowDatas(EnvParams.Country,"Currency");
  ExcelUtils.setExcelName(workBook, "ExchangeRate", true);
  if(ContryCurrency!="GBP")  
  ExchangeRate = ExcelUtils.getRowDatas(ContryCurrency,"Exchange Rate");
  else
  ExchangeRate = "1.00";
  if(ClientCurrency!="GBP")  
  var BaseCurrency = ExcelUtils.getRowDatas(ClientCurrency.OleValue.toString().trim(),"Exchange Rate");
  else
  BaseCurrency = "1.00";

  
 for(var i=1;i<=10;i++){
   Log.Message(sheetName);
ExcelUtils.setExcelName(workBook, sheetName, true);
var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
var Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
var Employee_Categories = ExcelUtils.getColumnDatas("Employee Categories_"+i,EnvParams.Opco)
var Employee_Number =  ExcelUtils.getColumnDatas("Employee Number_"+i,EnvParams.Opco)
var CostBase = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
var OutwardHSN = ExcelUtils.getColumnDatas("Outward HSN_"+i,EnvParams.Opco)
var InwardHSN = ExcelUtils.getColumnDatas("Inward HSN_"+i,EnvParams.Opco)

if((wCodeID!="")&&(wCodeID!=null)){

var AddBudget = "";
if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.isVisible()){
AddBudget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
}else{ 
AddBudget = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite;  
}
Log.Message(AddBudget.FullName)
Sys.HighlightObject(AddBudget);
var linest = false
for(var kk= 0;kk<AddBudget.ChildCount;kk++){
  if(AddBudget.Child(kk).isVisible()){ 
      if((AddBudget.Child(kk).Name.indexOf("SingleToolItemControl")!=-1)&&(AddBudget.Child(kk).Name.indexOf("4")!=-1)){
       AddBudget =  AddBudget.SWTObject("SingleToolItemControl", "", 4);
       linest = true;
       }
  }
}
  
if(linest){
WorkspaceUtils.waitForObj(AddBudget);
ReportUtils.logStep_Screenshot("");
AddBudget.Click(); 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}else{ 
var copy = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite2.SingleToolItemControl3;
WorkspaceUtils.waitForObj(AddBudget);
ReportUtils.logStep_Screenshot(""); 
copy.Click();
  

var removeZeroBudgetLine = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite2.SingleToolItemControl2;
WorkspaceUtils.waitForObj(removeZeroBudgetLine);
removeZeroBudgetLine.Click();
aqUtils.Delay(3000, "Jobs - Budget");
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim())!=-1){
var ApiButton = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
ApiButton.HoverMouse();
ReportUtils.logStep_Screenshot("");
ApiButton.Click();
}
aqUtils.Delay(3000, "Jobs - Budget");
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim())!=-1){
var ApiButton = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
ApiButton.HoverMouse();
ReportUtils.logStep_Screenshot("");
ApiButton.Click();
}
    
}
  
//-----Work Code Selection---------    

if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.isVisible()){
var Clientgrid = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
}else{ 
var Clientgrid = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid; 
}
Log.Message(Clientgrid.FullName)
//var Clientgrid = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var workcode;
linestatus = false;
workcode = Clientgrid.SWTObject("McValuePickerWidget", "");

if(wCodeID!=""){
addedlines = true;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
WorkspaceUtils.waitForObj(workcode);
workcode.Click();
WorkspaceUtils.SearchByValue(workcode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),wCodeID,"Work Code :"+wCodeID);
}else{ 
  ValidationUtils.verify(false,true,"WorkCode Needed to create JobBudget");
}
WorkspaceUtils.waitForObj(workcode);

    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Next Column");
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Next Column");
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Next Column");

//-----Internal Description---------
linestatus = false;
var External_Description;
 External_Description = Clientgrid.SWTObject("McTextWidget", "", 4);

    Sys.HighlightObject(External_Description);
    External_Description.Click();
    if(Desp!=""){
    External_Description.setText(Desp);
    ValidationUtils.verify(true,true,"External Description is Entered");
    }
    
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Next Column");

//-----Employee Categories if required---------
//var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
var EmpCat;
linestatus = false;
EmpCat = Clientgrid.SWTObject("McValuePickerWidget", "");

if(Employee_Categories!=""){
EmpCat.Click();
WorkspaceUtils.SearchByValue(EmpCat,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee Category").OleValue.toString().trim(),Employee_Categories,"Employee Category");
 }
         
Sys.Desktop.KeyDown(0x09); // Press Ctrl
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(1000, "Next Column");

    
//-----Employee Number if required---------    
var empno;
linestatus = false;
empno = Clientgrid.SWTObject("McValuePickerWidget", "");

Sys.HighlightObject(empno);
if(Employee_Number!=""){
empno.Click();
WorkspaceUtils.SearchByValue(empno,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),Employee_Number,"Employee Number");
     }
         
//    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Next Column");

    
//-----Quantity---------
linestatus = false;
var Quality = Clientgrid.SWTObject("McTextWidget", "", 2);

    Sys.HighlightObject(Quality);
    Quality.Click();
    if(Qly!=""){
    Quality.setText(Qly);
    ValidationUtils.verify(true,true,"Quality is Entered");
    }

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(2000, "Next Column");
    
//-----Cost Base Only for Amount---------

  for(var yy=0;yy<workActivity.length;yy++){ 
  if((workActivity[yy].indexOf(wCodeID)!=-1)&&((workActivity[yy].indexOf("Outlays")!=-1)||(workActivity[yy].indexOf("Desembolsos")!=-1))){ 
  wCodeID  = "T1001";
  break;
  }  
  }

if(wCodeID.indexOf("T")==-1){
var Cost_base;
Cost_base = Clientgrid.SWTObject("McTextWidget", "", 2);
linestatus = false;

    Sys.HighlightObject(Cost_base);
    Cost_base.Click();
    if(CostBase!=""){    
    Cost_base.setText(CostBase);
    ValidationUtils.verify(true,true,"Cost is Entered");
    }
    }
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    
//-----Cost Base Only for Time---------
    if(wCodeID.indexOf("T")>-1){
var Billing_Price;
 Billing_Price = Clientgrid.SWTObject("McTextWidget", "", 2);
linestatus = false;

    Sys.HighlightObject(Billing_Price);
    Billing_Price.Click();
    if(CostBase!=""){      
    Billing_Price.setText(CostBase);
    ValidationUtils.verify(true,true,"Cost is Entered");
    }
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
Outward_HSN = Clientgrid.SWTObject("McValuePickerWidget", "");
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
linestatus = false;
Invard_HSN = Clientgrid.SWTObject("McValuePickerWidget", "");

    Sys.HighlightObject(Invard_HSN);
    if(InwardHSN!=""){
    Invard_HSN.Click();
    WorkspaceUtils.SearchByValue(Invard_HSN,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 9").OleValue.toString().trim(),InwardHSN,"Inward HSN");
         }else{ 
    ValidationUtils.verify(false,true,"Inward HSN Needed to create JobBudget");
    }
  
}

var Save = "";
    if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.isVisible()){
    Save = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
    }else{ 
    Save = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
    }
    Log.Message(Save.FullName)
    Sys.HighlightObject(Save);
    Save.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    Save.Click();
    aqUtils.Delay(3000, "Saving added lines in Work Estimate");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    
/*
  // validation part
  var tableGrid = Clientgrid;
  var total_Cost_Base = tableGrid.getItem(RowCount).getText_2(8).OleValue.toString().trim();
  var Billing_Price_Curr = tableGrid.getItem(RowCount).getText_2(9).OleValue.toString().trim();
  var total_Billing_Price_Currency = tableGrid.getItem(RowCount).getText_2(10).OleValue.toString().trim();
  var Tax_code1 = tableGrid.getItem(RowCount).getText_2(15).OleValue.toString().trim();
  var Tax_code2 = tableGrid.getItem(RowCount).getText_2(16).OleValue.toString().trim();
  total_Cost_Base = total_Cost_Base.replace(/,/g, '');
  Billing_Price_Curr = Billing_Price_Curr.replace(/,/g, '');
  total_Billing_Price_Currency = total_Billing_Price_Currency.replace(/,/g, '');
  var t2 = "0.00"
  if(wCodeID.indexOf("T")>-1){
  var tcb = parseFloat(Qly)*parseFloat("0");
  t2 = parseFloat(CostBase);
  t2 = t2.toFixed(2);
  }else{ 
  var tcb = parseFloat(Qly)*parseFloat(CostBase);
  var t1= parseFloat(CostBase)/parseFloat(ExchangeRate); //Exchange Rate is Opco Currency
  t2 = parseFloat(t1)*parseFloat(BaseCurrency); //Base Currency is Client Currency
  t2 = t2.toFixed(2);
  }
  
  var tBPC = parseFloat(Billing_Price_Curr)*parseFloat(Qly);
  Log.Message(tBPC)
  tcb = tcb.toFixed(2);
  tBPC = tBPC.toFixed(2);

var lowerRange = parseFloat(t2)-parseFloat("5.00");
var higherRange = parseFloat(t2)+parseFloat("1.00");
Log.Message(lowerRange)
Log.Message(higherRange)
Log.Message(Billing_Price_Curr)
Log.Message(tBPC)
Log.Message(total_Billing_Price_Currency)

  if(tcb==total_Cost_Base)
  ValidationUtils.verify(true,true,"Total Cost Base is verified");
  else
  ValidationUtils.verify(false,true,"Total Cost Base is Not Matched ");
  
  if((lowerRange<Billing_Price_Curr)&&(higherRange>Billing_Price_Curr))
  ValidationUtils.verify(true,true,"Billing_Price_Curr is verified");
  else
  ValidationUtils.verify(false,true,"Billing_Price_Curr is Not Matched ");
  
  if(tBPC==total_Billing_Price_Currency)
  ValidationUtils.verify(true,true,"Total Billing Price Currency is verified");
  else
  ValidationUtils.verify(false,true,"Total Billing Price Currency is Not Matched ");
  
  if((Tax_code1=="")&&(Tax_code2==""))
  ValidationUtils.verify(false,true,"Tax Code 1 and Tax Code 2 is not Populated");
  if(Tax_code1!="")
  ValidationUtils.verify(true,true,"Tax Code 1 is populated");
  if(Tax_code2!="")
  ValidationUtils.verify(true,true,"Tax Code 2 is populated");
  
  TotalBudget = parseFloat(TotalBudget.toString()) + parseFloat(total_Billing_Price_Currency.toString());
  */
  
RowCount++;
} 
}
if(!addedlines)
ValidationUtils.verify(false,true,"WorkCode is not availble in to Create Budget");
else{ 

var Submit = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.SingleToolItemControl;
Log.Message(Submit.FullName)
WorkspaceUtils.waitForObj(Submit);
ReportUtils.logStep_Screenshot("");
Submit.Click();

ValidationUtils.verify(true,true,"Created Budget is Submitted");
TextUtils.writeLog("Working Estimate is Submitted"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var BarStat = true;
quteNumber = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6;  
for(var a=0;a<quteNumber.ChildCount;a++){ 
  Log.Message(quteNumber.Child(a).FullName)
  if((quteNumber.Child(a).isVisible())&&(quteNumber.Child(a).Name.indexOf("Composite")!=-1)&&(quteNumber.Child(a).Index==1)){ 
    quteNumber = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget.getText().OleValue.toString().trim();
    Log.Message(quteNumber);
    BarStat = false;
    break;
  }
}
if(BarStat){
quteNumber = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5;  
for(var a=0;a<quteNumber.ChildCount;a++){ 
  if((quteNumber.Child(a).isVisible())&&(quteNumber.Child(a).Name.indexOf("Composite")!=-1)&&(quteNumber.Child(a).Index==1)){ 
    quteNumber = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget.getText().OleValue.toString().trim();
    Log.Message(quteNumber);
    BarStat = false;
    break;
  }
}
}

if(BarStat){
quteNumber = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8;  
for(var a=0;a<quteNumber.ChildCount;a++){ 
  if((quteNumber.Child(a).isVisible())&&(quteNumber.Child(a).Name.indexOf("Composite")!=-1)&&(quteNumber.Child(a).Index==1)){ 
    quteNumber = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget.getText().OleValue.toString().trim();
    Log.Message(quteNumber);
    BarStat = false;
    break;
  }
}
}

Log.Message("quteNumber :"+quteNumber)

var BarStat = true;
var AprveBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("Composite", "", 6);  
for(var a=0;a<AprveBar.ChildCount;a++){ 
  Log.Message(AprveBar.Child(a).FullName)
  if((AprveBar.Child(a).isVisible())&&(AprveBar.Child(a).Name.indexOf("PTabItemPanel")!=-1)&&(AprveBar.Child(a).Index==3)){ 
    AprveBar = AprveBar.Child(a).SWTObject("TabControl", "")
    Log.Message(AprveBar.FullName);
    BarStat = false;
    break;
  }
}
if(BarStat){
AprveBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7;  
for(var a=0;a<AprveBar.ChildCount;a++){ 
  if((AprveBar.Child(a).isVisible())&&(AprveBar.Child(a).Name.indexOf("PTabItemPanel")!=-1)&&(AprveBar.Child(a).Index==3)){ 
    AprveBar = AprveBar.Child(a).SWTObject("TabControl", "");
    Log.Message(AprveBar.FullName);
    break;
  }
}
}

if(BarStat){
AprveBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6;  
for(var a=0;a<AprveBar.ChildCount;a++){ 
  if((AprveBar.Child(a).isVisible())&&(AprveBar.Child(a).Name.indexOf("PTabItemPanel")!=-1)&&(AprveBar.Child(a).Index==3)){ 
    AprveBar = AprveBar.Child(a).SWTObject("TabControl", "");
    Log.Message(AprveBar.FullName);
    break;
  }
}
}

AprveBar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var AprveAction ="";
//All_Approver = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
var AprveAction = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite;
for(var a=0;a<AprveAction.ChildCount;a++){ 
  if(AprveAction.Child(a).isVisible()){ 
    Log.Message(AprveAction.Child(a).FullName);
    var SubAd = AprveAction.Child(a);
    for(var b=0;b<SubAd.ChildCount;b++){ 
      if((SubAd.Child(b).isVisible())&&(SubAd.Child(b).Name.indexOf("Composite")!=-1)&&(SubAd.Child(b).Index==2)){ 
    AprveAction = AprveAction.Child(a).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    Log.Message(AprveAction.FullName);
    break;
    }
    }
  }
}
    Sys.HighlightObject(AprveAction)
    AprveAction.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
} 
linestatus = false;
if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.isVisible()){
var Approval_table = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
    Sys.HighlightObject(Approval_table)
    }
else{
var Approval_table = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
    Sys.HighlightObject(Approval_table)
    }
    Sys.HighlightObject(Approval_table)
    for(var z=0;z<Approval_table.getItemCount();z++){ 
       approvers="";
       approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
       Approve_Level[y] = comapany+"*"+jobNumber+"*"+approvers;
       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
       y++;
    }

linestatus = false;
if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.isVisible()){
var ApprovalTableBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl;
    Sys.HighlightObject(ApprovalTableBar)
    }
else{
var ApprovalTableBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.PTabItemPanel.TabControl;
    Sys.HighlightObject(ApprovalTableBar)
    }
Sys.HighlightObject(ApprovalTableBar)
ApprovalTableBar.Click(); 
if(ImageRepository.ImageSet.Forward.Exists()){
ImageRepository.ImageSet.Forward.Click();// GL
}
TextUtils.writeLog(Approve_Level.length+" Levels of Approvals for Created Budget");

CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
//sheetName = "JobBudgetCreation";

if(OpCo2[2]==Project_manager){
  
level = 1;

var BarStat = true;
var Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6;  
for(var a=0;a<Approve.ChildCount;a++){ 
  Log.Message(Approve.Child(a).FullName)
  if((Approve.Child(a).isVisible())&&(Approve.Child(a).Name.indexOf("Composite")!=-1)&&(Approve.Child(a).Index==1)){ 
    Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl;
    Log.Message(Approve.FullName);
    BarStat = false;
    break;
  }
}
if(BarStat){
Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9;  
for(var a=0;a<Approve.ChildCount;a++){ 
  if((Approve.Child(a).isVisible())&&(Approve.Child(a).Name.indexOf("Composite")!=-1)&&(Approve.Child(a).Index==1)){ 
    Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.SingleToolItemControl;
    Log.Message(Approve.FullName);
    break;
  }
}
}

Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
WorkspaceUtils.waitForObj(Approve);
ReportUtils.logStep_Screenshot("");
Approve.Click();
ValidationUtils.verify(true,true,"Levels 0 has  Approved the Created Budget");
TextUtils.writeLog("Levels 0 has  Approved the Created Budget");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var BarStat = true;
var ApvPerson = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6;  
for(var a=0;a<ApvPerson.ChildCount;a++){ 
  Log.Message(ApvPerson.Child(a).FullName)
  if((ApvPerson.Child(a).isVisible())&&(ApvPerson.Child(a).Name.indexOf("Composite")!=-1)&&(ApvPerson.Child(a).Index==1)){ 
    ApvPerson = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McTextWidget;;
    Log.Message(ApvPerson.FullName);
    BarStat = false;
    break;
  }
}
if(BarStat){
ApvPerson = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9;  
for(var a=0;a<ApvPerson.ChildCount;a++){ 
  if((ApvPerson.Child(a).isVisible())&&(ApvPerson.Child(a).Name.indexOf("Composite")!=-1)&&(ApvPerson.Child(a).Index==1)){ 
    ApvPerson = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget
    Log.Message(ApvPerson.FullName);
    break;
  }
}
}

var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while ((ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

  if(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1){
  ValidationUtils.verify(true,true,"Job is Approved by :"+loginPer)
  TextUtils.writeLog("Job is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Job is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Job is Approved by :"+loginPer+ "But its Not Reflected")
  }
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(ComId+" - "+JobNo +" - Approver :"+loginPer);
}

if(ApproveInfo.length == 1){
//  try{
TextUtils.writeLog("Budget is created for :"+jobNumber);
TextUtils.writeLog("Revision : "+quteNumber);
//ExcelUtils.setExcelName(workBook,"CreateQuote", true);
//ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Revision","CreateQuote",quteNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Working Estimate_"+Jobnum,EnvParams.Opco,"Data Management",jobNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Budget Revision No_"+Jobnum,EnvParams.Opco,"Data Management",quteNumber);

//}
//catch(err){ 
//  Log.Message(err);
//}
}

////Approve.Click();
//var Add_Visible8 = true;
//while(Add_Visible8){
//aqUtils.Delay(2000, Indicator.Text);
////Delay(2000);
//if(Approve.isEnabled_2){
//Add_Visible8 = false;
//Sys.HighlightObject(Approve);
//Approve.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//Approve.Click();
//ValidationUtils.verify(true,true,"Levels 0 has  Approved the Created Budget");
//TextUtils.writeLog("Levels 0 has  Approved the Created Budget");
//}
//}
}

}

}

 
function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}  
    
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
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
//     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, "Agency Users", true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

//    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, "SSC Users", true);
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
//WorkspaceUtils.closeAllWorkspaces();
}

    
  

function todo(lvl){ 
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
if(((temp.indexOf("Approve Job Budget by Type (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo por tipo (")!=-1))&&(temp1.length==2)){ 
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
if(((temp.indexOf("Approve Job Budget by Type (Substitute) (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo por tipo (sustituto) (")!=-1))&&(temp1.length==3)){ 
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
if(((temp.indexOf("Approve Job Budget (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo (")!=-1))&&(temp1.length==2)){ 
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
if(((temp.indexOf("Approve Job Budget (Substitute) (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo (sustituto) (")!=-1))&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job Budget (Substitute) from To-Dos List");
var listPass = false;   
  }
} 
  }

}




function aprvBudget(ComId,JobNo,userNmae){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "");
waitForObj(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
Sys.HighlightObject(table);

if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).Visible){

}else{ 
var showFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
waitForObj(table);
Sys.HighlightObject(showFilter);
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.Click();
}

 var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
    SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
    SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).
    SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
    SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
    SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
   
//    aqUtils.Delay(2000, Indicator.Text);
//    Delay(2000);
    
    companyFilter.forceFocus();
    companyFilter.setVisible(true);
    companyFilter.ClickM();
waitForObj(companyFilter);
Sys.HighlightObject(companyFilter);
companyFilter.HoverMouse();
companyFilter.HoverMouse();
    table.Child(0).setText("^a[BS]");
    table.Child(0).setText(ComId);
    aqUtils.Delay(1000, "Moving to Job Name");;
//    Delay(2000);
    table.Child(0).Keys("[Tab][Tab]");

    var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    
    job.ClickM();
    table.Child(2).forceFocus();
    table.Child(2).setVisible(true);
    table.Child(2).setText("^a[BS]");
    table.Child(2).setText(JobNo);
    aqUtils.Delay(3000, "Reading Data in table");;
//    Delay(3000);
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
    
    if(table.getItemCount()>0){
//    Log.Message("Created Job is listed in table")
TextUtils.writeLog("Created JobBudget is listed in Approval list");
ReportUtils.logStep_Screenshot("");
    closeFilter.Click();
//    Delay(8000);
    
//    var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
//    Budget.Click();
//    Delay(2000);
//    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
//    Sys.HighlightObject(show_budget);
//    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

//    var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
WorkspaceUtils.waitForObj(Budget);
    Budget.Click();
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    ref.Refresh();
    aqUtils.Delay(5000, "Waiting for McClumpSashForm");
//    Delay(5000);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  } 
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").Index==1)
    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==1)
    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    Sys.HighlightObject(show_budget);
    

    show_budget.Keys("Working Estimate");
    aqUtils.Delay(2000, Indicator.Text);
//    Delay(7000);
    var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 9);

Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
  Approve.HoverMouse();
ReportUtils.logStep_Screenshot("");
  Approve.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  } 
var ApvPerson = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while ((ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

    if(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1){
  ValidationUtils.verify(true,true,"Job is Approved by :"+userNmae)
  TextUtils.writeLog("Job is Approved by :"+userNmae); 
  }else{ 
  TextUtils.writeLog("Job is Approved by :"+userNmae+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Job is Approved by :"+userNmae+ "But its Not Reflected")
  }
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
}

if((ApproveInfo.length -1)== level){
  try{
TextUtils.writeLog("Budget is created for :"+jobNumber);
TextUtils.writeLog("Revision : "+quteNumber);
//ExcelUtils.setExcelName(workBook,"CreateQuote", true);
//ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Revision","CreateQuote",quteNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Working Estimate_"+Jobnum,EnvParams.Opco,"Data Management",jobNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Budget Revision No_"+Jobnum,EnvParams.Opco,"Data Management",quteNumber);
}catch(err){ 
  Log.Message(err);
}
}

    }

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
//Log.Message("Notepad :"+notepadPath)
return TextUtils.readDetails(notepadPath,"Job Number");
//Log.Message( readDetails("C:\\Users\\674087\\Documents\\TestComplete 14 Projects\\After Stuart Discussion\\WppRegression_v12.50\\WppRegPack\\RegressionLogs\\TESTAPAC\\Regression\\China\\1307-Sudler China(MDS)_Client Billable.txt","Job Number") );
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
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
//Log.Message(compStatus);
return compStatus
}

function Clientgrid(){ 
var mainRoot = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");      
    var workcode;
    var linestatus = false;
    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("3")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}


    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("2")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}


    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("4")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}


    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("5")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}
    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
 var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("7")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}
return workcode;
}