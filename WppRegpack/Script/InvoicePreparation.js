//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "InvoicePreparation";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var percentage,jobNumber = "";

//Main Function
function InvoicePreparation(){ 
TextUtils.writeLog("Create General Journal Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Biller",EnvParams.Opco);
if((Project_manager=="")||(Project_manager==null))
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if((Project_manager=="")||(Project_manager==null))
ValidationUtils.verify(false,true,"Login Credentials required for anyone of Agency - Biller or Agency - Finance,");

Log.Message(Project_manager);
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
sheetName = "InvoicePreparation";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
percentage,jobNumber = "";
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Invoice Preparation started::"+STIME);

getDetails();
gotoMenu();
gotoInvoicing();

}

function getDetails(){ 
sheetName ="InvoicePreparation";  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed for Invoice Preparation");
  

  

  ExcelUtils.setExcelName(workBook, sheetName, true);
  percentage = ExcelUtils.getRowDatas("Percentage",EnvParams.Opco)
  if((percentage=="")||(percentage==null))
  ValidationUtils.verify(false,true,"Percentage is needed for Invoice Preparation");
}


function gotoMenu(){ 
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
Client_Managt.ClickItem("|Jobs");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Jobs");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}

function gotoInvoicing(){ 
var allJobs = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
allJobs.Click();

//var labels = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget.Label;
var labels = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf("Now showing")!=-1)){
    labels = labels.Child(i);
    break;
  }
}

WorkspaceUtils.waitForObj(labels);
var i=0;
while((labels.getText().OleValue.toString().trim().indexOf("results")!=-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf("results")!=-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

  var table = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JobTable;
  var firstcell = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JobTable.CompanyNo;
  var closeFilter = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter;
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
  
  var job = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.JobTable.JobNumber;
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);

var i=0;
while((labels.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf("results")==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
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
  ReportUtils.logStep("INFO", "Job is listed in table to for Invoice Preparation");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy for Invoice Preparation"); 
  closeFilter.Click();

/*
  var Budget = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
  WorkspaceUtils.waitForObj(Budget);
  Budget.Click();
  
  var showBudget = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  WorkspaceUtils.waitForObj(showBudget);
  showBudget.Keys("Client Approved Estimate");
  WorkspaceUtils.waitForObj(showBudget);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
//  var fullBudget = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl;
  var fullBudget = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.TabFolderPanel.TabControl
  WorkspaceUtils.waitForObj(fullBudget);
  fullBudget.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
//  var table = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var table = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
  WorkspaceUtils.waitForObj(table);
  var Estimate = [];
  for(var i=0;i<table.getItemCount();i++){ 
  Estimate[i] = table.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+table.getItem(i).getText_2(6).OleValue.toString().trim();
  }
  
  */
  
  var Invoicing = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Invoicing;
//  var Invoicing = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.TabFolderPanel.TabControl
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();
  

  
  var iPreparation = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.invoicePreparation;
  WorkspaceUtils.waitForObj(iPreparation);
  iPreparation.Click();
  
  
  var CreateIPreparation = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.EditableBar;
  WorkspaceUtils.waitForObj(CreateIPreparation);
  for(var i=0;i<CreateIPreparation.ChildCount;i++){ 
    if((CreateIPreparation.Child(i).isVisible())&&(CreateIPreparation.Child(i).toolTipText=="Prepare Invoice")){
      CreateIPreparation.Child(i).Click();
      break;
    }
  }
  
  var BasicOfInvoice = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  BasicOfInvoice.Keys("Budget");

var i=0;
while((BasicOfInvoice.getText()!="Budget")&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  BasicOfInvoice.Refresh();
}
if(BasicOfInvoice.getText()!="Budget"){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

var next = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite2.Composite.Composite.Button;
next.Click();

var Percentage = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.Composite.Percentage;
WorkspaceUtils.waitForObj(Percentage);
Percentage.Click();
Percentage.setText(percentage);

var prepare = Aliases.Maconomy.PrepareInvoice.Composite.Composite.Composite2.Composite.Button;
WorkspaceUtils.waitForObj(prepare);
prepare.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Jobs - Invoice Preparation"){
var Okay = Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Invoice Preparation").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
Okay.Click();
}

var invoiceTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(invoiceTable);
for(var i=0;i<invoiceTable.getItemCount()-1;i++){ 
 
//for(var j=0;j<Estimate.length;j++) { 
//  Esplit = Estimate[j].split("*");
//  
//  if(Esplit[0]==invoiceTable.getItem(i).getText_2(0).OleValue.toString().trim()){
    
  var temp = invoiceTable.getItem(i).getText_2(1).OleValue.toString().trim();
  temp = parseFloat(temp).toFixed(2);;
//  if(percentage.indexOf(".")!=-1){ 
    percentage = parseFloat(percentage.toString()).toFixed(2);
//  }
  Log.Message("temp :"+temp)
  Log.Message("percentage :"+percentage)
  if(temp!=percentage){ 
    ValidationUtils.verify(true,false,"Percentage missMatch with Datasheet");
  }
  
//  var temp = invoiceTable.getItem(i).getText_2(2).OleValue.toString().trim();
//  Esplit[1] = parseFloat(Esplit[1].toString())/100;
//  Esplit[1] = parseFloat(Esplit[1]*percentage).toFixed(2);
//  
//  if(Esplit[1]!=temp){ 
//  Log.Message("Esplit :"+Esplit[1])
//  Log.Message("temp :"+temp)   
//  ValidationUtils.verify(true,false,"Quatity missMatch with repective percentage");
//  }
//  break;
//  }
  
//}

}

var BudgetAmount = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite2.NewTotal;
BudgetAmount = BudgetAmount.getText();



  var Submit = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.EditableBar;
  WorkspaceUtils.waitForObj(Submit);
  for(var i=0;i<Submit.ChildCount;i++){ 
    if((Submit.Child(i).isVisible())&&(Submit.Child(i).toolTipText=="Submit")){
      WorkspaceUtils.waitForObj(Submit.Child(i));
      Submit.Child(i).Click();
      break;
    }
  }
  
var removeLatestDrafft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.SingleToolItemControl;
var i=0;
while((!removeLatestDrafft.Visible)&&(removeLatestDrafft.getText()=="Remove Latest Draft")&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  removeLatestDrafft.Refresh();
}
if(removeLatestDrafft.getText()!="Remove Latest Draft"){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

WorkspaceUtils.waitForObj(removeLatestDrafft);
  
var draftInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(draftInvoice);
draftInvoice.Click();

var draftNo = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(draftNo);
draftNo.Keys("[Tab][Tab][Tab]");
var billiablePrice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(billiablePrice);
billiablePrice.Click();
billiablePrice.setText(BudgetAmount);

var DraftTable = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var flag=false;
  for(var v=0;v<DraftTable.getItemCount();v++){ 
    if(DraftTable.getItem(v).getText_2(3).OleValue.toString().trim()==BudgetAmount){ 
      flag=true;
      break;
    }
    else{ 
      DraftTable.Keys("[Down]");
    }
  }
 ValidationUtils.verify(true,flag,"Prepared Invoice is available to submit Draft")
  if(flag){
var CloseFilter = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;  
CloseFilter.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
var SubmitDraft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;
//                  Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite
                  
  WorkspaceUtils.waitForObj(SubmitDraft);
  for(var i=0;i<SubmitDraft.ChildCount;i++){ 
    if((SubmitDraft.Child(i).isVisible())&&(SubmitDraft.Child(i).toolTipText=="Submit Draft")){
      WorkspaceUtils.waitForObj(SubmitDraft.Child(i));
      SubmitDraft.Child(i).Click();
      break;
    }
  }
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
    
 var PrintDraft = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.EditableBar;
  WorkspaceUtils.waitForObj(PrintDraft);
  for(var i=0;i<PrintDraft.ChildCount;i++){ 
    if((PrintDraft.Child(i).isVisible())&&(PrintDraft.Child(i).toolTipText=="Print Draft")){
      WorkspaceUtils.waitForObj(PrintDraft.Child(i));
      PrintDraft.Child(i).Click();
      break;
    }
  } 
  
TextUtils.writeLog("Print Draft is Clicked");
//    aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Invoice Editing")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x41);
    
    if(ImageRepository.PDF.ChooseFolder.Exists())
    ImageRepository.PDF.ChooseFolder.Click();
    
    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
    aqUtils.Delay(2000, Indicator.Text);
    SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
saveAs.Click();
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Order Confirmation is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);
    
  
  }
  }

}
