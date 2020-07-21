﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateAnAccrualForAGroupOfJobs";
var Language = "";
Indicator.Show();
  
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
//var comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile;
var JobNoFrom,JobNoTo,WorkCode,EntryDate,NoForAccrual,PoNumber,JobNumber;


//getting data from datasheet
function getDetails(){
//Log.Message("excelName :"+workBook);
//Log.Message("sheet :"+sheetName);
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(EnvParams.Opco)
JobNoFrom = ExcelUtils.getRowDatas("JobNumberFrom",EnvParams.Opco)
if((JobNoFrom==null)||(JobNoFrom=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to Create a Job");
}
Log.Message(JobNoFrom)


JobNoTo = ExcelUtils.getRowDatas("JobNumberTo",EnvParams.Opco)
if((JobNoTo==null)||(JobNoTo=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to Create a Job");
}
Log.Message(JobNoTo)

EntryDate = ExcelUtils.getRowDatas("EntryDate",EnvParams.Opco)
if((EntryDate==null)||(EntryDate=="")){ 
ValidationUtils.verify(false,true,"EntryDate is Needed to Create a Job");
}
Log.Message(EntryDate)

JobNumber = ExcelUtils.getRowDatas("JobNumber",EnvParams.Opco)
if((JobNumber==null)||(JobNumber=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to Create a Job");
}
Log.Message(JobNumber)


PoNumber = ExcelUtils.getRowDatas("PoNumber",EnvParams.Opco)
if((PoNumber==null)||(PoNumber=="")){ 
ValidationUtils.verify(false,true,"PoNo is Needed to Create a Job");
}
Log.Message(PoNumber)

WorkCode = ExcelUtils.getRowDatas("WorkCode",EnvParams.Opco)
if((WorkCode==null)||(WorkCode=="")){ 
ValidationUtils.verify(false,true,"WorkCode is Needed to Create a Job");
}
Log.Message(WorkCode)


NoForAccrual = ExcelUtils.getRowDatas("NoForAccrual",EnvParams.Opco)
if((NoForAccrual==null)||(NoForAccrual=="")){ 
ValidationUtils.verify(false,true,"NoForAccrual Number is Needed to Create a Job");
}
Log.Message(NoForAccrual)


//PoNumber_1 = ExcelUtils.getRowDatas("PoNumber_1",EnvParams.Opco)
//if((PoNumber_1==null)||(PoNumber_1=="")){ 
//ValidationUtils.verify(false,true,"PoNoFrom is Needed to Create a Job");
//}
//Log.Message(PoNumber_1)
//
//
//WorkCode_1 = ExcelUtils.getRowDatas("WorkCode_0",EnvParams.Opco)
//if((WorkCode_1==null)||(WorkCode_1=="")){ 
//ValidationUtils.verify(false,true,"WorkCodeTo is Needed to Create a Job");
//}
//Log.Message(WorkCode_1)
//
//
//
//NoForAccrual_1 = ExcelUtils.getRowDatas("NoForAccrual_1",EnvParams.Opco)
//if((NoForAccrual_1==null)||(NoForAccrual_1=="")){ 
//ValidationUtils.verify(false,true,"NoForAccrual Number is Needed to Create a Job");
//}
//Log.Message(NoForAccrual_1)


//ExcelUtils.setExcelName(workBook, "Data Management", true);
//Product = ReadExcelSheet("Global Product Number",EnvParams.Opco,"Data Management");
//if((Product=="")||(Product==null)){
//ExcelUtils.setExcelName(workBook, sheetName, true);
//Product = ExcelUtils.getRowDatas("Product",EnvParams.Opco)
//}
//if((Product==null)||(Product=="")){ 
//ValidationUtils.verify(false,true,"Product Number is Needed to Create a Job");
//}

//ExcelUtils.setExcelName(workBook, sheetName, true);
//Job_name= ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
//if((Job_name==null)||(Job_name=="")){ 
//ValidationUtils.verify(false,true,"Job Name is Needed to Create a Job");
//}
Dlang= ExcelUtils.getRowDatas("Language",EnvParams.Opco)

//BFC= ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)

//pTerm= ExcelUtils.getRowDatas("Payment Terms",EnvParams.Opco)
ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}




function GoToAccruals() {
  
var Accrualtab =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.AccrualTab;
  WorkspaceUtils.waitForObj(Accrualtab);
Accrualtab.Click();
 
var BatchJobAccrualtab =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.BatchAccrualJobtAB;
   WorkspaceUtils.waitForObj(BatchJobAccrualtab);
BatchJobAccrualtab.Click();
 
//var JobNoTextBox = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.JobSearchField;


var jobNoFrom =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.JobNofROM;

var JobNoTO =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.JobNoTo;


//jobNoFrom.setText(JobNoFrom);
//JobNoTO.setText(JobNoTo);

 if(JobNoFrom!=""){
  jobNoFrom.Click();
  
  WorkspaceUtils.SearchByValues_all_Col_2(jobNoFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),JobNoFrom,"Job Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
 // WorkspaceUtils.SearchByValue(jobNoFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),JobNoFrom,"Job Number");
    }
 else{ 
    ValidationUtils.verify(false,true,"JobNoFrom is Needed to Create a Payment Selection");
  }
  
  
   if(JobNoTo!=""){
  JobNoTO.Click();
  WorkspaceUtils.SearchByValues_all_Col_2(JobNoTO,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),JobNoTo,"Job Number",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  //WorkspaceUtils.SearchByValue(JobNoTO,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),JobNoTo,"Job No");
    }
 else{ 
    ValidationUtils.verify(false,true,"JobNoTo is Needed to Create a Payment Selection");
  }

 aqUtils.Delay(500);
  var showlines =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPlainCheckboxView.ShowLines;
 aqUtils.Delay(500);
  var includeFullyAccured =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McPlainCheckboxView.IncludeFullyAccrued;
  aqUtils.Delay(500);
  
  //----------De-Select CheckBox-------------
  if(!showlines.getSelection()){ 
  showlines.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showlines.Click();
  ReportUtils.logStep("INFO", "showlines is UnChecked");
    Log.Message("showlines is UnChecked")
    checkmark = true;
  }
   aqUtils.Delay(500);
  if(includeFullyAccured.getSelection()){ 
includeFullyAccured.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  includeFullyAccured.Click();
  ReportUtils.logStep("INFO", "includeFullyAccured is UnChecked");
//    Log.Message("Blanket_invoice is UnChecked")
    checkmark = true;
  }
  
   aqUtils.Delay(500);
  var purchaseorderNoFromField =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.POnoFrom

  var purchaseorderNoToField = Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.POnoTo;

  var workCodeFrom = Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.WorkCodeFrom;

  var workCodeTo = Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.WorkCodeTo;


 aqUtils.Delay(5000);
  
      var savejob =  Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SaveJob;
    
 //Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.savejobButton;
 if(savejob.isEnabled())
 {
 savejob.Click();
  }
  aqUtils.Delay(3000, "Waiting for purchaseOrderTable load");
  
  var purchaseOrderTable =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
  WorkspaceUtils.waitForObj(purchaseOrderTable);
 
  //purchaseOrderTable.Click();
  
    var flag=false;
  
   for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
  
    if(purchaseOrderTable.getItem(v).getText_2(5).OleValue.toString().trim()==(WorkCode)&&(purchaseOrderTable.getItem(v).getText_2(1).OleValue.toString().trim()==(PoNumber)&&(purchaseOrderTable.getItem(v).getText_2(0).OleValue.toString().trim()==(JobNumber)))){ 

    purchaseOrderTable.Keys("[Tab][Tab]");
    aqUtils.Delay(100);
    purchaseOrderTable.Keys(" ");
    aqUtils.Delay(1000);
    

    purchaseOrderTable.Keys("[Tab]");
    purchaseOrderTable.Keys(NoForAccrual);  
    
    aqUtils.Delay(500);   
    purchaseOrderTable.Keys("[Tab]");  
    aqUtils.Delay(500);

    purchaseOrderTable.Keys(EntryDate);  
    aqUtils.Delay(500);

  var savePOLine =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SavePO;
  
// Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SavePOLine
  
  savePOLine.Click();
  aqUtils.Delay(3000);
//    var MarkForAccrual =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.markBatchAccruals;
//  MarkForAccrual.Click();
    
  aqUtils.Delay(1000);
   flag=true;
      break;
      
    }
    else{ 
      purchaseOrderTable.Keys("[Down]");
    }
  }
  
  if(flag){
  ValidationUtils.verify(flag,true,"Purchase Order Line with Work Code is available in system");
  ValidationUtils.verify(true,true,"Batch Accrual is Successful");
  }
  else{
     ValidationUtils.verify(flag,true,"Purchase Order Line with Work Code is not available in system");
  ValidationUtils.verify(flag,true,"Batch Accrual is not Successful");
  }
  
//  var savejob1 =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SaveJob;
//  
//if(savejob1.isEnabled()){
//savejob1.Click();
//   }
if(Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.Visible){
 var CreateAccruals =Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.CreateAccruals
 }
 else {
   var CreateAccruals = Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.CreateAccruals;
 }

 Sys.HighlightObject(CreateAccruals)
CreateAccruals.Click();
   
  }
  
   
function test()
{
  
  var purchaseOrderTable =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
  
  purchaseOrderTable.Click();
  purchaseOrderTable.Keys("[PageUp][PageUp]");
  
 var flag=false; 
  for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
    
    if(purchaseOrderTable.getItem(v).getText_2(5).OleValue.toString().trim()==("E1022")){ 

purchaseOrderTable.Keys("[Tab][Tab]");
purchaseOrderTable.Keys("[Space][Tab]");
purchaseOrderTable.k
  
    
   flag=true;
      break;
    }
    else{ 
      purchaseOrderTable.Keys("[Down]");
    }
  }
  Log.Message(flag);
//     var purchaseorderNoFromField=Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget
//     purchaseorderNoFromField.setText("1707109766");
}

//Validating Created Job is available


//Go To Job from Menu
function goToJobMenuItem(){

//aqUtils.Delay(5000, Indicator.Text);
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

//if(ImageRepository.ImageSet.Jobs1.Exists()){
//ImageRepository.ImageSet.Jobs1.Click();// GL
//}
//
//
//if(ImageRepository.ImageSet3.Jobs.Exists()){
// ImageRepository.ImageSet3.Jobs.Click();// GL
//}
//else if(ImageRepository.ImageSet.Job.Exists()){
//ImageRepository.ImageSet.Job.Click();
//}
//else{
////ImageRepository.ImageSet.Jobs1.Click();
// ImageRepository.ImageSet3.Jobs.Click();
//}

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//Delay(3000);
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());


}

}



ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}






//Main Function
function CreateAnAccrualGroupOfJobs() {
TextUtils.writeLog("Job Creation Started"); 
Indicator.PushText("waiting for window to open");
//aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)

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

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateAnAccrualForAGroupOfJobs";
Language = "";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
//comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile ="";
JobNoFrom,JobNoTo,WorkCode,EntryDate,NoForAccrual,PoNumber,JobNumber="";

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);
aqUtils.Delay(3000, Indicator.Text);
getDetails();
goToJobMenuItem();   
GoToAccruals();
WorkspaceUtils.closeAllWorkspaces();

//
}

//function getExcel(rowidentifier,column) { 
//excelData =[];  
////Log.Message(" ");
////Log.Message(excelName)
////Log.Message(workBook);
////Log.Message(sheetName);
//var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
//var id =0;
//var colsList = [];
// var temp ="";
////Log.Message(rowidentifier);
//     while (!DDT.CurrentDriver.EOF()) {
////Log.Message(xlDriver.Value(0).toString().trim())
////Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
//       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
////       Log.Message("Row Identifier is Matched");
//        try{
//         temp = temp+xlDriver.Value(column).toString().trim();
//         }
//        catch(e){
//        temp = "";
//        }
////      Log.Message(temp);
//      break;
//      }
//
//    xlDriver.Next();
//     }
//     
//     if(temp.indexOf(",")!=-1){
//     var excelData =  temp.split(",");
////     Log.Message(excelData);
////     for(var i=0;i<comma_separator.length;i++){ 
////       
////     }
//       
//     }else if(temp.length>0){ 
//      excelData[0] = temp;
////       excelData[0] = temp.substring(0, temp.indexOf("-"));
////       excelData[1] = temp.substring(temp.indexOf("-")+1)
//     }
//     
//     DDT.CloseDriver(xlDriver.Name);
// for(var i=0;i<excelData.length;i++)
//// Log.Message(excelData[i]);
//     return excelData;
//  
//}



function getExcelData(rowidentifier,column) { 
var temp = ""

var excelData = [];
//Log.Message(workBook+":")
//Log.Message(sheetName+":")
//Log.Message(rowidentifier+":")
//Log.Message(column+":")
ExcelUtils.setExcelName(workBook, sheetName, true);
temp = ExcelUtils.getRowDatas(rowidentifier,column);
//Log.Message(temp);
//temp = temp.OleValue.toString().trim();

/*
  var app = Sys.OleObject("Excel.Application");
  var curArrayVals = [];  
  var book = app.Workbooks.Open(workBook);
  var sheet = book.Sheets.Item(sheetName);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;

  var arrays={};
  var idx =0;
  var col =0;
  var row = 0;
  for(var k = 1; k<=columnCount;k++){
  if(sheet.Cells.Item(1, k).Text.toString().trim().toUpperCase()==column.toUpperCase()){
  col = k;
  }
  }
  var rowStatus = false;
  for(var k = 1; k<=rowCount;k++){
  if(sheet.Cells.Item(k, 1).Text.toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
  row = k;
  rowStatus = true;
  }
  }
  if(rowStatus){ 
   temp = sheet.Cells.Item(row,  col).Text;

  }
  
 app.Quit();
*/
 
 if(temp.indexOf(",")!=-1){ 
//       Log.Message(temp)
      excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     

// for(var i=0;i<excelData.length;i++)
// Log.Message(" :"+excelData[i]);

 return excelData;
}

function getExcelData_Company(rowidentifier,column) { 
var excelData =[];  
var temp ="";
ExcelUtils.setExcelName(workBook, sheetName, true);
temp = ExcelUtils.getRowDatas(rowidentifier,column);
//temp = temp.OleValue.toString().trim();

/*
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
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
     
DDT.CloseDriver(xlDriver.Name);
*/
     
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
     
//     DDT.CloseDriver(xlDriver.Name);

// for(var i=0;i<excelData.length;i++)
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



