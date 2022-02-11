//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

/**
 * This script to print Intercompany Invoice for Expenses
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Created Date :07/05/2021
*/

//Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "IC Invoice";
var Intercompany_OpCo,Expense_Number = "";


//Main Function
function IC_Invoice_Printing() {
  
TextUtils.writeLog("IC Invoice Printing Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

ExcelUtils.setExcelName(workBook, "IC Invoice", true);
Intercompany_OpCo = ExcelUtils.getRowDatas("InterCompany OpCo",EnvParams.Opco)
if((Intercompany_OpCo==null)||(Intercompany_OpCo=="")){ 
ValidationUtils.verify(false,true,"InterCompany OpCo is Needed to Create a InterCompany Client");
}
Log.Message(Intercompany_OpCo)

//Checking Login to execute Import Budget Template script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Intercompany","Username");
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "IC Invoice";
STIME = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "IC Invoice started::"+STIME);
Expense_Number = "";

getDetails();
goTo_GL_Transaction();
goTo_Invoice();



}



function getDetails(){ 
  
sheetName = "IC Invoice";
ExcelUtils.setExcelName(workBook, "Data Management", true);
Expense_Number = ReadExcelSheet("InterCompany Expense Number",EnvParams.Opco,"Data Management");
if((Expense_Number=="")||(Expense_Number==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Expense_Number = ExcelUtils.getRowDatas("Expense Number",EnvParams.Opco)
} 
if((Expense_Number=="")||(Expense_Number==null)){
 ValidationUtils.verify(true,false,"Expense Number is need to print Intercompany Invoice");
}

Log.Message(Expense_Number)
}


function goTo_GL_Transaction(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.GendralLedger.Exists()){
ImageRepository.ImageSet.GendralLedger.Click();// GL
}
else if(ImageRepository.ImageSet.GendralLedger1.Exists()){
ImageRepository.ImageSet.GendralLedger1.Click();
}
else{
ImageRepository.ImageSet.GendralLedger2.Click();
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Transactions from General Ledger Menu");
TextUtils.writeLog("Entering into GL Transactions from General Ledger Menu");
}




function goTo_Invoice(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var intercompany_Invoice = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(intercompany_Invoice);
intercompany_Invoice.Click()

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Company_Num = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Company_Num.Click();
Company_Num.Keys(EnvParams.Opco);


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;



var flag = false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()== EnvParams.Opco){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  
  
if(flag){ 
  var CloseFilter = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(CloseFilter);
  CloseFilter.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var Show_Lines = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McPlainCheckboxView.Button;
Sys.HighlightObject(Show_Lines);
if(!Show_Lines.getSelection()){
Show_Lines.Click();
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var Intercompany_Entries = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(Intercompany_Entries);
Intercompany_Entries.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

var Entries = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var flag = false;
  for(var v=0;v<Entries.getItemCount();v++){ 
    if(Entries.getItem(v).getText_2(4).OleValue.toString().trim()== Expense_Number){ 
      flag=true;
      Entries.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
      break;
    }
    else{ 
      Entries.Keys("[Down]");
    }
  }
  
  if(flag){ 
    var CheckBox = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
    Sys.HighlightObject(CheckBox);
    CheckBox.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(3000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var Save = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(Save);
    Save.Click();
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(3000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var print_InterCompany_Invoice = Aliases.Maconomy.InterCompany_Invoicing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
    Sys.HighlightObject(print_InterCompany_Invoice);
    print_InterCompany_Invoice.Click();


    aqUtils.Delay(5000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(5000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      
  
  
    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
    aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice"+"*", 1).WndCaption.indexOf("Invoice")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x12); 
    Sys.Desktop.KeyUp(0x46); //Alt
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
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");


ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF IC Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  

    
  }

}



}