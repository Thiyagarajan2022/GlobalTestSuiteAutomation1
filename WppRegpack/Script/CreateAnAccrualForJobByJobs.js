﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateAnAccrualJobByJob";
var Language = "";
  Indicator.Show();
  
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var JobNo,WorkCodeFrom,WorkCodeTo,EntryDate,NoForAccrual,PoNoFrom,PoNoTo,NoOfPoLinesForAccrual ;


//getting data from datasheet
function getDetails(){
//Log.Message("excelName :"+workBook);
//Log.Message("sheet :"+sheetName);
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(EnvParams.Opco)

JobNo = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
if((JobNo==null)||(JobNo=="")){ 
ValidationUtils.verify(false,true,"Job Number is Needed to CreateAnAccrualJobByJob");
}

Log.Message(JobNo)

WorkCodeFrom = ExcelUtils.getRowDatas("WorkCodeFrom",EnvParams.Opco)
if((WorkCodeFrom==null)||(WorkCodeFrom=="")){ 
ValidationUtils.verify(false,true,"WorkCodeFrom is Needed to CreateAnAccrualJobByJob");
}

Log.Message(WorkCodeFrom)

WorkCodeTo = ExcelUtils.getRowDatas("WorkCodeTo",EnvParams.Opco)
if((WorkCodeTo==null)||(WorkCodeTo=="")){ 
ValidationUtils.verify(false,true,"WorkCodeTo is Needed to CreateAnAccrualJobByJob");
}
Log.Message(WorkCodeTo)

EntryDate = ExcelUtils.getRowDatas("EntryDate",EnvParams.Opco)
if((EntryDate==null)||(EntryDate=="")){ 
ValidationUtils.verify(false,true,"EntryDate is Needed to CreateAnAccrualJobByJob");
}
Log.Message(EntryDate)

NoForAccrual = ExcelUtils.getRowDatas("NoForAccrual",EnvParams.Opco)
if((NoForAccrual==null)||(NoForAccrual=="")){ 
ValidationUtils.verify(false,true,"NoForAccrual Number is Needed to CreateAnAccrualJobByJob");
}
Log.Message(NoForAccrual)

PoNoFrom = ExcelUtils.getRowDatas("PoNoFrom",EnvParams.Opco)
if((PoNoFrom==null)||(PoNoFrom=="")){ 
ValidationUtils.verify(false,true,"PoNoFrom is Needed to CreateAnAccrualJobByJob");
}
Log.Message(PoNoFrom)

PoNoTo = ExcelUtils.getRowDatas("PoNoTo",EnvParams.Opco)
if((PoNoFrom==null)||(PoNoFrom=="")){ 
ValidationUtils.verify(false,true,"PoNoTo is Needed to CreateAnAccrualJobByJob");
}
Log.Message(PoNoTo)

NoOfPoLinesForAccrual =ExcelUtils.getRowDatas("NoOfPoLinesForAccrual",EnvParams.Opco)
if((NoOfPoLinesForAccrual==null)||(NoOfPoLinesForAccrual=="")){ 
ValidationUtils.verify(false,true,"PoNoTo is Needed to Create a Job");
}
Log.Message(NoOfPoLinesForAccrual)

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
//Dlang= ExcelUtils.getRowDatas("Language",EnvParams.Opco)

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
 Accrualtab.Click();
 
var JobNoTextBox = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.JobSearchField;

JobNoTextBox.setText(JobNo);

var table = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//WorkspaceUtils.waitForObj(jobAccrualTable);


var labels= Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.LabelOneOfOneResult;
//Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.
//McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf("Now showing")!=-1)){
      labels = labels.Child(i);
      break;
    }
  }

  WorkspaceUtils.waitForObj(labels);

  var i=0;
  while((labels.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=600)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf("results")==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}


  aqUtils.Delay(2000, "Reading Table Data in Job List");
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(JobNo)){ 

      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(0).OleValue.toString().trim());

  var closefilter = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.Closefilter;
  closefilter.Click();
  
  var jobAccrualPannel = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel;
  jobAccrualPannel.Click();
  jobAccrualPannel.MouseWheel(-200);
  
  
  var showlines =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite.McPlainCheckboxView.ShowLinesCheckBox;
  var includeFullyAccured =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite2.McPlainCheckboxView.inclueFullyAccured;
  
  //----------De-Select CheckBox-------------
  if(!showlines.getSelection()){ 
  showlines.HoverMouse();
ReportUtils.logStep_Screenshot("");
  showlines.Click();
  ReportUtils.logStep("INFO", "showlines is UnChecked");
    Log.Message("showlines is UnChecked")
    checkmark = true;
  }
  
  if(includeFullyAccured.getSelection()){ 
includeFullyAccured.HoverMouse();
  ReportUtils.logStep_Screenshot("");
  includeFullyAccured.Click();
  ReportUtils.logStep("INFO", "includeFullyAccured is UnChecked");
//    Log.Message("Blanket_invoice is UnChecked")
    checkmark = true;
  }
  
  
  
  var purchaseorderNoFromField = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget
  var purchaseorderNoToField = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite4.PurchaseorderToField;
  
  var purchaseorderlineNoField = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite3.PruchaseOrderFrom;
  
  var workCodeFrom = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite5.WorkCodeField;
  var workCodeTo = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite5.WorkCodeTo;

  
    
  
//purchaseorderNoFromField.setText("1707109766");
//  
//var purchaseorderToValue = "";
//
//if((purchaseorderToValue==null)||(purchaseorderToValue=="")){ 
//purchaseorderNoToField.setText("1707109766");  


//var NoOfPurchaseOrderLinesForAccrual =1
Log.Message("NoOfPoLinesForAccrual"+NoOfPoLinesForAccrual);

for(var i=1;i<=NoOfPoLinesForAccrual;i++){ 
  
//purchaseorderNoFromField.setText(PoNoFrom)
//purchaseorderNoToField.setText(PoNoTo);  



 if(PoNoFrom!=""){
  purchaseorderNoFromField.Click();
  WorkspaceUtils.SearchByValue(purchaseorderNoFromField,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),PoNoFrom,"Purchase Order");
    }
 else{ 
    ValidationUtils.verify(false,true,"PaymentNo is Needed to Create a Payment Selection");
  }
  
  
  if(PoNoTo!=""){
  purchaseorderNoToField.Click();
  WorkspaceUtils.SearchByValue(purchaseorderNoToField,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order").OleValue.toString().trim(),PoNoTo,"Purchase Order");
    }
 else{ 
    ValidationUtils.verify(false,true,"PaymentNo is Needed to Create a Payment Selection");
  }
  
  

//  workCodeFrom.setText(WorkCodeFrom);  
//  workCodeTo.setText(WorkCodeTo);  
//  
  
    if(WorkCodeFrom!=""){
  workCodeFrom.Click();
  WorkspaceUtils.SearchByValue(WorkCodeFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),WorkCodeFrom,"Work Code");
    }
 else{ 
    ValidationUtils.verify(false,true,"PaymentNo is Needed to Create a Payment Selection");
  }
  
  
    if(WorkCodeTo!=""){
  workCodeTo.Click();
  WorkspaceUtils.SearchByValue(workCodeTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Work Code").OleValue.toString().trim(),WorkCodeTo,"Work Code");
    }
 else{ 
    ValidationUtils.verify(false,true,"PaymentNo is Needed to Create a Payment Selection");
  }
  
  
  
      var savejob =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.savejobButton;
  savejob.Click();
  
  aqUtils.Delay(3000, "Waiting for purchaseOrderTable load");
  
  var purchaseOrderTable =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
 
  purchaseOrderTable.Click();
  
    var flag=false;
  
   for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
  
    if(purchaseOrderTable.getItem(v).getText_2(5).OleValue.toString().trim()==(WorkCodeTo)&&(purchaseOrderTable.getItem(v).getText_2(0).OleValue.toString().trim()==(PoNoFrom))){ 

      flag=true;
    Sys.Keys("[Tab][Tab][Tab]");
    
    aqUtils.Delay(500);
    
    var noForAccrual =
Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable.NoForrAccrual    
//Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable.NoForAccrual;
   Sys.HighlightObject(noForAccrual)
 noForAccrual.setText(NoForAccrual);  
    
  aqUtils.Delay(500);   
   noForAccrual.Keys("[Tab]");  
 aqUtils.Delay(500);
  var entryDate = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable.EntryDate;
     Sys.HighlightObject(entryDate)
  entryDate.setText(EntryDate);  
    aqUtils.Delay(500);

  var savePOLine = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SavePOLine
  
  savePOLine.Click();
  aqUtils.Delay(3000);
    var MarkForAccrual =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.MarkForAccrual;
  MarkForAccrual.Click();
    
  aqUtils.Delay(1000);
 
      break;
      
    }
    else{ 
      purchaseOrderTable.Keys("[Down]");
    }
  }
  
  ValidationUtils.verify(flag,true,"Purchase Order Line with Work Code is available in system");
  ValidationUtils.verify(true,true,"Purchase Order :"+table.getItem(v).getText_2(0).OleValue.toString().trim());
  
   
  //Empty the Fields
//  purchaseorderNoFromField.setText("");
//  purchaseorderNoToField.setText("");  
//  workCodeFrom.setText("");     
//  workCodeTo.setText("");
  
  
  var savejob =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.savejobButton;
  savejob.Click();
  
 
  
  
  }
  
   var CreateAccruals =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.CreateAccrual;
  
   CreateAccruals.Click();
   
   
  }
  
//  var purchaseOrderTable =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
//  
//  
//  
//  for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
//    var flag=false;
//    if(purchaseOrderTable.getItem(v).getText_2(5).OleValue.toString().trim()==("E1003")){ 
//purchaseOrderTable.getItem(v).get
//    
//   flag=true;
//    //  break;
//    }
//    else{ 
//      purchaseOrderTable.Keys("[Down]");
//    }
//  }
  
  
   
  


function test()
{
  

var purchaseOrderTable =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
 
  purchaseOrderTable.Click();
  Sys.HighlightObject(purchaseOrderTable)
    var flag=false;
  
   for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
  Log.Message("purchaseOrderTable.getItemCount()"+purchaseOrderTable.getItemCount())
    if(purchaseOrderTable.getItem(v).getText_2(5).OleValue.toString().trim()==("E1003")&&(purchaseOrderTable.getItem(v).getText_2(0).OleValue.toString().trim()==("1707109766"))){ 
 flag=true;
    Sys.Keys("[Tab][Tab][Tab]");
    
    aqUtils.Delay(500);
    
    var noForAccrual =
Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable.NoForrAccrual    
//Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable.NoForAccrual;
   Sys.HighlightObject(noForAccrual)
 noForAccrual.setText(NoForAccrual);  
    
  aqUtils.Delay(500);   
   noForAccrual.Keys("[Tab]");  
 aqUtils.Delay(500);
  var entryDate = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable.EntryDate;
     Sys.HighlightObject(entryDate)
  entryDate.setText("7/1/2020");  
    aqUtils.Delay(500);

  var savePOLine = Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SavePOLine
  
  savePOLine.Click();
  aqUtils.Delay(3000);
    var MarkForAccrual =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.MarkForAccrual;
  MarkForAccrual.Click();
    
  aqUtils.Delay(1000);
  
      break;
      
    }
    else{ 
      purchaseOrderTable.Keys("[Down]");
    }
  }
  
  
  
//  var purchaseOrderTable =Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.PurchaseOrderTable;
//  
//  purchaseOrderTable.Click();
//  purchaseOrderTable.Keys("[PageUp][PageUp]");
//  
// var flag=false; 
//  for(var v=0;v<purchaseOrderTable.getItemCount();v++){ 
//    
//    if(purchaseOrderTable.getItem(v).getText_2(5).OleValue.toString().trim()==("E1022")){ 
//
//purchaseOrderTable.Keys("[Tab][Tab]");
//purchaseOrderTable.Keys("[Space][Tab]");
//purchaseOrderTable.k
//  
//    
//   flag=true;
//      break;
//    }
//    else{ 
//      purchaseOrderTable.Keys("[Down]");
//    }
//  }
//  Log.Message(flag);
////     var purchaseorderNoFromField=Aliases.Maconomy.JobAccruals.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.AccrualPanel.Accurlapanel.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget
////     purchaseorderNoFromField.setText("1707109766");
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
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());
//ReportUtils.logStep_Screenshot();
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Administration").OleValue.toString().trim());
  Client_Managt.ClickItem("|Job Administration");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|Job Administration");

}

}



//var mainlist = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
//var main;
//for(var id=0;id<mainlist;id++){
//main = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//if(main.Child(id).isVisible())
//if(main.Child(id).ChildCount==1)
//if(main.Child(id).Child(0).Name.indexOf("Composite")!=-1){
//
//var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
//  var Client_Managt;
////Log.Message(childCC)
//for(var i=1;i<=childCC;i++){ 
//Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
//if(Client_Managt.isVisible()){ 
//Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
//ReportUtils.logStep_Screenshot("");
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
//}
//}
//}
//
//}
//Delay(5000); 
ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}






//Main Function
function CreateAnAccrualJobByJob() {
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
sheetName = "CreateAnAccrualJobByJob";
Language = "";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
//comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile ="";
JobNo,WorkCodeFrom,WorkCodeTo,EntryDate,NoForAccrual,PoNoFrom,PoNoTo,NoOfPoLinesForAccrual ="";

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
//createAJob();   
//GoToJob();
//WorkspaceUtils.closeAllWorkspaces();
//aqTestCase.End();

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



