﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "InvoiceLookups";
var Project_manager="";
var STIME = "";
var InvoiceNo = "";

//Main Function
function InvoiceLookUps(){ 
TextUtils.writeLog("Create Invoice Lookups Started"); 
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
sheetName = "InvoiceLookups";
STIME = "";
InvoiceNo = "";
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Invoice Invoice started::"+STIME);

getDetails();
gotoMenu();
Lookups();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
sheetName ="InvoiceLookups";  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  InvoiceNo = ReadExcelSheet("Client Invoice No",EnvParams.Opco,"Data Management");
  if((InvoiceNo=="")||(InvoiceNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  InvoiceNo = ExcelUtils.getRowDatas("Invoice No",EnvParams.Opco)
  }
  if((InvoiceNo=="")||(InvoiceNo==null))
  ValidationUtils.verify(false,true,"Invoice No is needed for Invoice LookUps");
  
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
Client_Managt.ClickItem("|Job Invoices");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Job Invoices");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Job Invoices Menu");
TextUtils.writeLog("Entering into Jobs from Job Invoices Menu");
}

function Lookups(){ 
  
  var labels = Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
  WorkspaceUtils.waitForObj(labels);
  for(var i=0;i<labels.ChildCount;i++){ 
    if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf("Now showing")!=-1)){
      labels = labels.Child(i);
      break;
    }
  }

  WorkspaceUtils.waitForObj(labels);

  
  var table = Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  var closeFilter = Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.setText(InvoiceNo);
  WorkspaceUtils.waitForObj(firstcell);
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
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==InvoiceNo){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Invoice is listed in table to for Invoice Lookups");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Invoice("+InvoiceNo+") is available in maconommy for Invoice Lookups"); 
  closeFilter.Click();
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var PrintCopy = Aliases.Maconomy.InvoiceLookUps.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite;
  WorkspaceUtils.waitForObj(PrintCopy);
  for(var i=0;i<PrintCopy.ChildCount;i++){ 
  if((PrintCopy.Child(i).isVisible())&&(PrintCopy.Child(i).toolTipText=="Print Copy")){
  PrintCopy.Child(i).HoverMouse(); 
  ReportUtils.logStep_Screenshot("");
  PrintCopy.Child(i).Click();
  break;
  }
  }
  
  TextUtils.writeLog("Print Copy is Clicked");
  var SaveTitle = "";
  var sFolder = "";
  var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Job Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
  if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Job Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Job Invoice")!=-1){
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
  else{ 
  var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
  WorkspaceUtils.waitForObj(window);
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x73); //F4
  Sys.Desktop.KeyUp(0x12); //Alt
  Sys.Desktop.KeyUp(0x73); //F4
  aqUtils.Delay(2000, Indicator.Text);
  Sys.HighlightObject(pdf)
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x46); //F
  Sys.Desktop.KeyDown(0x41); //A 
  Sys.Desktop.KeyUp(0x12); 
  Sys.Desktop.KeyUp(0x46); //Alt
  Sys.Desktop.KeyUp(0x41);
  }
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
//  var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//  saveAs.Click();
  var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
  aqUtils.Delay(2000, Indicator.Text);
  Sys.HighlightObject(pdf);
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x46); //F
  Sys.Desktop.KeyDown(0x58); //X 
  Sys.Desktop.KeyUp(0x46); //Alt
  Sys.Desktop.KeyUp(0x12);     
  Sys.Desktop.KeyUp(0x58);
  }
  ValidationUtils.verify(true,true,"Print Job Invoice is Clicked and PDF is Saved");
  Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
  ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
  aqUtils.Delay(4000, Indicator.Text);
   

  
  }
}