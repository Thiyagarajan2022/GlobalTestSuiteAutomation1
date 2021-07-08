﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart




/** 
 * This script Create Working Estimate for Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Modified Date :11/18/2020
 */
 
Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName ="InterCompany Budget";
var Intercompany_OpCo = "";
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
var Project_manager = "";




// Main Function
function createBudget(){ 
  
TextUtils.writeLog("Job Budget Creation Started"); 
Indicator.PushText("waiting for window to open");


//Getting Language from EnvParamaters.xlsx
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

ExcelUtils.setExcelName(workBook, "InterCompany Budget", true);
Intercompany_OpCo = ExcelUtils.getColumnDatas("InterCompany OpCo",EnvParams.Opco)
if((Intercompany_OpCo==null)||(Intercompany_OpCo=="")){ 
ValidationUtils.verify(false,true,"InterCompany OpCo is Needed to Create a InterCompany Client");
}
Log.Message(Intercompany_OpCo)
//Checking Login to execute Create Budget
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}

//Re-Intialize Variable
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

try{

getDetails();
goToJobMenuItem();
goToBudget();
sheetName ="InterCompany Budget";
addingBudgetLines();
closeAllWorkspaces();
  
// Approving Created Budget
for(var i=level;i<ApproveInfo.length;i++){
level = i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");

Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
aprvBudget(temp[0],temp[1],temp[2]);
}

}
  catch(err){
    Log.Message(err);
  }
  
//Close all Open Workspace in Maconomy
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



// Getting Details from Datasheet to Create Budget
function getDetails(){ 

comapany = EnvParams.Opco
sheetName ="InterCompany Budget";
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("InterCompany Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  sheetName ="InterCompany Budget";
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Create Budget");

var CodeStatus = true;
var Country = EnvParams.Country;

 for(var i=1;i<=10;i++){
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



// Navigating to Job from Jobs Menu
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


// Navigating to Budget Screen
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


// Selecting Working Estimate in show Budget
var show_budget = "";
  var BarStat = true;
  var ChildCount = 0;
  var Add = [];
  var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
   Sys.Process("Maconomy").Refresh();
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
    var show_budget = Parent.SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);

    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);
    WorkspaceUtils.waitForObj(show_budget);
    show_budget.HoverMouse();
    show_budget.HoverMouse();
    show_budget.HoverMouse();


// Getting Client Currency for Calculating Tax
  var BarStat = true;
  var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
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
    ClientCurrency = Parent.SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
    Log.Message(ClientCurrency.FullName)
    Sys.HighlightObject(ClientCurrency);


    Log.Message(ClientCurrency.FullName)
    WorkspaceUtils.waitForObj(ClientCurrency);
    ClientCurrency = ClientCurrency.getText();
    
    
// Selcting Working Estimate
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



// Adding Each Budget Lines for Job
function addingBudgetLines(){ 
  
  // Selecting Full Budget
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

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }



//Clicking Add line Icon

var ChildCount = 0;
var Add = [];
var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
for(var ip=0;ip<Parent.ChildCount;ip++){ 
var PChild = Parent.Child(ip);
if((PChild.isVisible()) && (PChild.ChildCount==3)){
for(var jp=0;jp<PChild.ChildCount;jp++){
if((PChild.Child(jp).isVisible()) && (PChild.Child(jp).JavaClassName=="Composite") && (PChild.Child(jp).Index==2)) { 
Add[ChildCount] = PChild.Child(jp);
ChildCount++;
}
}
}
}
     
     var AddBudget = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       AddBudget = Add[ip];
     }     
     }
     
     Sys.HighlightObject(AddBudget)
     Log.Message(AddBudget.FullName)
     AddBudget = AddBudget.SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
     Sys.HighlightObject(AddBudget)

 
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
 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  } 
if(linest){
WorkspaceUtils.waitForObj(AddBudget);
ReportUtils.logStep_Screenshot("");
AddBudget.Click(); 
}else{ 
  

// Copy from Budget Template
var copy = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite2.SingleToolItemControl3;
WorkspaceUtils.waitForObj(copy);
ReportUtils.logStep_Screenshot(""); 
copy.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
                          
// Remove Zero Budget Lines
    var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
            if((PChild.isVisible()) && (PChild.ChildCount==1)){
         Add[ChildCount] = PChild;
         ChildCount++;
     }
     }
     
     var removeZeroBudgetLine = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       removeZeroBudgetLine = Add[ip];
     }     
     }
     
     Sys.HighlightObject(removeZeroBudgetLine)
     Log.Message(removeZeroBudgetLine.FullName)
     
     if(removeZeroBudgetLine.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible()){ 
     removeZeroBudgetLine = removeZeroBudgetLine.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 14); 
     }else{ 
      removeZeroBudgetLine = removeZeroBudgetLine.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 12); 
     }
     
     Sys.HighlightObject(removeZeroBudgetLine)
WorkspaceUtils.waitForObj(removeZeroBudgetLine);
removeZeroBudgetLine.Click();
aqUtils.Delay(5000, "Jobs - Budget");




//Handling PoP - Ups

 aqUtils.Delay(10000, "Reporting in HTML about Notification"); 
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim())
  var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim(), 2000);
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

  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim())
  var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Budget").OleValue.toString().trim(), 2000);
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

// Adding New Line Icon
 aqUtils.Delay(10000, "Reporting in HTML about Notification"); 
    var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
            if((PChild.isVisible()) && (PChild.ChildCount==3)){
     for(var jp=0;jp<PChild.ChildCount;jp++){
       if((PChild.Child(jp).isVisible()) && (PChild.Child(jp).JavaClassName=="Composite") && (PChild.Child(jp).Index==2)) { 
         Add[ChildCount] = PChild.Child(jp);
         ChildCount++;
       }
     }
     }
     }
     
     var AddBudget = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       AddBudget = Add[ip];
     }     
     }
     
     Sys.HighlightObject(AddBudget)
     Log.Message(AddBudget.FullName)
     AddBudget = AddBudget.SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);
     Sys.HighlightObject(AddBudget)

 
Log.Message(AddBudget.FullName)
Sys.HighlightObject(AddBudget);
var linest = false
Log.Message(AddBudget.ChildCount)
       AddBudget =  AddBudget.SWTObject("SingleToolItemControl", "", 4);
       Sys.HighlightObject(AddBudget)
       AddBudget.Click();
       linest = true;
   
}
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }


//-----Work Code Selection---------    

// Finding Grid Address
var Clientgrid = ClientgridTable();
Log.Message(Clientgrid.FullName)


//Select WorkCode
var workcode;
linestatus = false;
workcode = Clientgrid.SWTObject("McValuePickerWidget", "");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
if(wCodeID!=""){
addedlines = true;
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
  // Tax validation part
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

// Submit Budget
if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.isVisible())
var Submit = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.SingleToolItemControl;
else
var Submit = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);

Log.Message(Submit.FullName)
WorkspaceUtils.waitForObj(Submit);
ReportUtils.logStep_Screenshot("");
Submit.Click();

ValidationUtils.verify(true,true,"Created Budget is Submitted");
TextUtils.writeLog("Working Estimate is Submitted"); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}


// Capturing Revision Number
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


// Moving to sliding pane
var BarStat = true;
var AprveBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite;
for(var a=0;a<AprveBar.ChildCount;a++){ 
  if(AprveBar.Child(a).isVisible()){ 
    var temp = AprveBar.Child(a);
    Log.Message(temp.FullName)
    for(var b=0;b<temp.ChildCount;b++){ 
      if((temp.Child(b).isVisible())&&(temp.Child(b).Name.indexOf("PTabItemPanel")!=-1)){ 
       AprveBar =  temp.Child(b);
       Log.Message(AprveBar.FullName)
       break;
      
      }
    }

  }

}
AprveBar.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ImageRepository.ImageSet.Maximize.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

// Selecting All Approver
var AprveAction ="";
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
       Approve_Level[y] = Intercompany_OpCo +"*"+jobNumber+"*"+approvers;
       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
       y++;
    }

    
//Closing Sliding Pane
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
sheetName = "InterCompany Budget";

Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
if(OpCo2[2]==Project_manager){
  
level = 1;
// Approving Budget for Level 0
var BarStat = true;
var Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6;  
for(var a=0;a<Approve.ChildCount;a++){ 
  Log.Message(Approve.Child(a).FullName)
  if((Approve.Child(a).isVisible())&&(Approve.Child(a).Name.indexOf("Composite")!=-1)&&(Approve.Child(a).Index==1)){ 
    if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.isVisible())
    Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl;
    else
    Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
    Log.Message(Approve.FullName);
    BarStat = false;
    break;
  }
}
if(BarStat){
Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9;  
for(var a=0;a<Approve.ChildCount;a++){ 
  if((Approve.Child(a).isVisible())&&(Approve.Child(a).Name.indexOf("Composite")!=-1)&&(Approve.Child(a).Index==1)){ 
    if(Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.isVisible())
    Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.SingleToolItemControl;
    else
    Approve = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7)
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

// Moving to Sliding Pane to Validate whether its Approved or NOT
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


  ValidationUtils.verify(true,true,"Job is Approved by :"+loginPer)
  TextUtils.writeLog("Job is Approved by :"+loginPer); 

  var BarStat = true;
var AprveBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite;
for(var a=0;a<AprveBar.ChildCount;a++){ 
  if(AprveBar.Child(a).isVisible()){ 
    var temp = AprveBar.Child(a);
    Log.Message(temp.FullName)
    for(var b=0;b<temp.ChildCount;b++){ 
      if((temp.Child(b).isVisible())&&(temp.Child(b).Name.indexOf("PTabItemPanel")!=-1)){ 
       AprveBar =  temp.Child(b);
       Log.Message(AprveBar.FullName)
       break;
      
      }
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

    Sys.HighlightObject(Approval_table)
    
    if(Approval_table.getItem(0).getText_2(8).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,true,"Created Budget is Approved in Level : 0")
    }else{ 
      ValidationUtils.verify(true,false,"Created Budget is NOT Approved in Level :0")
    }
    
    if(Approval_table.getItem(0).getText_2(9).OleValue.toString().trim()==loginPer){
      ValidationUtils.verify(true,true,"Created Budget is Approved by :"+loginPer)
    }else{ 
      ValidationUtils.verify(true,false,"Created Budget is NOT Approved in Level :"+loginPer)
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

  
  
  
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(ComId+" - "+JobNo +" - Approver :"+loginPer);
}


// Storing Data in Datasheet
if(ApproveInfo.length == 1){
TextUtils.writeLog("Budget is created for :"+jobNumber);
TextUtils.writeLog("Revision : "+quteNumber);
ExcelUtils.setExcelName(workBook,"InterCompany Quote", true);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Revision","CreateQuote",quteNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("InterCompany Working Estimate",EnvParams.Opco,"Data Management",jobNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("InterCompany Budget Revision No",EnvParams.Opco,"Data Management",quteNumber);
}


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
    

// Finding Approver userName to Approve Budget
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
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(Intercompany_OpCo+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],Intercompany_OpCo);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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

    
  
// Finding Created Budget from To-Do's List
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

aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

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


    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    ref.Refresh();
    aqUtils.Delay(5000, "Waiting for McClumpSashForm");
//    Delay(5000);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, "Waiting for McClumpSashForm");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  
  // Selecting Working Estimate in Show Budget
    
      var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
   Sys.Process("Maconomy").Refresh();
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
    var show_budget = Parent.SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    Log.Message(show_budget.FullName)
    Sys.HighlightObject(show_budget);

    show_budget.Keys("Working Estimate");
    aqUtils.Delay(2000, Indicator.Text);


    var ChildCount = 0;
    var Add = [];

   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
            if((PChild.isVisible()) && (PChild.ChildCount==1)){
         Add[ChildCount] = PChild;
         ChildCount++;
     }
     }
     
     var Approe = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       Approe = Add[ip];
     }     
     }
     
     Sys.HighlightObject(Approe)
     Log.Message(Approe.FullName)
     Approe = Approe.SWTObject("Composite", "").SWTObject("PTabFolder", "");
     Sys.HighlightObject(Approe)
     
if(Approe.SWTObject("Composite", "", 2).isVisible())
 Approe = Approe.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 9);
else
 Approe = Approe.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
Sys.HighlightObject(Approe)

Sys.HighlightObject(Approe)
if(Approe.isEnabled()){ 
  Approe.HoverMouse();
ReportUtils.logStep_Screenshot("");
  Approe.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  } 

var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
  ValidationUtils.verify(true,true,"Job is Approved by :"+userNmae)
  TextUtils.writeLog("Job is Approved by :"+userNmae); 

// Validating Approved Status is Sliding Pane
  
  var AprveBar = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite;
for(var a=0;a<AprveBar.ChildCount;a++){ 
  if(AprveBar.Child(a).isVisible()){ 
    var temp = AprveBar.Child(a);
    Log.Message(temp.FullName)
    for(var b=0;b<temp.ChildCount;b++){ 
      if((temp.Child(b).isVisible())&&(temp.Child(b).Name.indexOf("PTabItemPanel")!=-1)){ 
       AprveBar =  temp.Child(b);
       Log.Message(AprveBar.FullName)
       break;
      
      }
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

var Approval_table = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;

    Sys.HighlightObject(Approval_table)
    
    if(Approval_table.getItem(level).getText_2(8).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,true,"Created Budget is Approved in Level :"+level)
    }else{ 
      ValidationUtils.verify(true,false,"Created Budget is NOT Approved in Level :"+level)
    }
    
    if(Approval_table.getItem(level).getText_2(9).OleValue.toString().trim()==loginPer){
      ValidationUtils.verify(true,true,"Created Budget is Approved by :"+loginPer)
    }else{ 
      ValidationUtils.verify(true,false,"Created Budget is NOT Approved in Level :"+loginPer)
    }
    

linestatus = false;

var ApprovalTableBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
Sys.HighlightObject(ApprovalTableBar)

Sys.HighlightObject(ApprovalTableBar)
ApprovalTableBar.Click(); 
if(ImageRepository.ImageSet.Forward.Exists()){
ImageRepository.ImageSet.Forward.Click();// GL
}

  
  
  
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
}


// Storing Data in Datasheet
if((ApproveInfo.length -1)== level){
TextUtils.writeLog("Budget is created for :"+jobNumber);
TextUtils.writeLog("Revision : 1");
ExcelUtils.setExcelName(workBook,"CreateQuote", true);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Revision","CreateQuote",quteNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Working Estimate",EnvParams.Opco,"Data Management",jobNumber);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Budget Revision No",EnvParams.Opco,"Data Management",quteNumber);
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

function ClientgridTable(){ 
  
    var ChildCount = 0;
    var Add = [];
    var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
    Sys.Process("Maconomy").Refresh();  
    for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
     if((PChild.isVisible()) && (PChild.ChildCount==3)){
       Log.Message(PChild.Name)
       for(var jp=0;jp<PChild.ChildCount;jp++){ 
         var CChild = PChild.Child(jp);
            if((CChild.isVisible()) && (CChild.JavaClassName=="Composite") && (CChild.Index==2)){
            Add[ChildCount] = CChild;
            ChildCount++;
            }
     }
     }
     }

     var grid = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       grid = Add[ip];
     }     
     }
     
     Log.Message(grid.FullName);
     Sys.HighlightObject(grid)
     grid = grid.SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
     Sys.HighlightObject(grid)

return grid;
}