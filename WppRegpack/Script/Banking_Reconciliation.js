//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


/**
 * This script Create Credit Note for Invoice
 * @author  : Muthu Kumar M
 * @version : 1.0
 * Created Date :04/07/2021
*/


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Banking Reconciliation";

//Global Variable
var Transaction_No,Account_No ,Statment_Date, Statment_No = "";

//Main Function
function Bank_Reconciliation() {
  
TextUtils.writeLog("Bank Reconciliation Creation Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Bank Reconciliation script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Treasury","Username");
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "Banking Reconciliation";
Transaction_No,Account_No ,Statment_Date, Statment_No = "";

getDetails();
aqUtils.Delay(5000, Indicator.Text);
goToBankReconciliation();
Search_Account_Number();
Reprint_Previous_Reconciliations();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}



function goToBankReconciliation(){ 
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.Banking.Exists()){
 ImageRepository.ImageSet.Banking.Click();// GL
}
else if(ImageRepository.ImageSet.Banking1.Exists()){
ImageRepository.ImageSet.Banking1.Click();
}
else{
ImageRepository.ImageSet.Banking2.Click();
}

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
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Reconciliations").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Reconciliations").OleValue.toString().trim());
}
}
ReportUtils.logStep("INFO", "Moved to Banking Reconciliations from WorkSpace Client");
TextUtils.writeLog("Entering into Banking Reconciliations from WorkSpace Client");

}



function getDetails(){ 
  


  ExcelUtils.setExcelName(workBook, "Data Management", true);
  Transaction_No = ExcelUtils.getRowDatas("Single Payment Trans No",EnvParams.Opco);
  if(Transaction_No == null)
  Transaction_No = ExcelUtils.getRowDatas("Multiple Payment Trans No",EnvParams.Opco);
  if(Transaction_No == null)
  Transaction_No = ExcelUtils.getRowDatas("Foreign Payment Trans No",EnvParams.Opco);
  
  Account_No = ExcelUtils.getRowDatas("Account Number",EnvParams.Opco);
  
  if((Account_No=="") && (Transaction_No=="") ){
  ExcelUtils.setExcelName(workBook, "Banking Reconciliation", true);
  Account_No = ExcelUtils.getRowDatas("Account Number",EnvParams.Opco);
  Transaction_No = ExcelUtils.getRowDatas("Transaction Number",EnvParams.Opco); 
  } 
   
  ExcelUtils.setExcelName(workBook, "Banking Reconciliation", true);
  Statment_Date = ExcelUtils.getRowDatas("Statement Date",EnvParams.Opco);
  if((Statment_Date=="") && (Statment_Date=="") ){
    ValidationUtils.verify(true,false,"Statement Date is needed to create Banking Reconciliation")
  }
  
  Statment_No = ExcelUtils.getRowDatas("Statement No",EnvParams.Opco);
  if((Statment_No=="") && (Statment_No=="") ){
    ValidationUtils.verify(true,false,"Statement Number is needed to create Banking Reconciliation")
  }
    
}

function Search_Account_Number(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Account_Number = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Account_Number.Click();
Account_Number.Keys(Account_No);
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

Account_Number.Keys("[Tab][Tab][Tab][Tab][Tab]");
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Company_No = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Company_No.Click();
Company_No.Keys(EnvParams.Opco);

aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table)
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 

    if((table.getItem(v).getText_2(0).OleValue.toString().trim()==Account_No) && (table.getItem(v).getText_2(5).OleValue.toString().trim()==EnvParams.Opco) ){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  
  if(flag){ 
    var CloseFilter = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(CloseFilter);
    CloseFilter.Click();
    
    aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);
   
   var Line = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
   Sys.HighlightObject(Line);
   Line.Click();
    aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   var Loop = true
   var table = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
   Sys.HighlightObject(table)
     var flag=false;
      for(var v=0;v<table.getItemCount()-1;v++){ 

        if(table.getItem(v).getText_2(1).OleValue.toString().trim()== Transaction_No){ 
          flag=true;
        // Selecting Transaction Number and Marking To-Reconcile Check Box
        if(Loop){
        var Date = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
        Sys.HighlightObject(Date);
        Date.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
        aqUtils.Delay(5000, Indicator.Text);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
        Loop = false;
        }
        
        Log.Message(Transaction_No)
        Log.Message(table.getItem(v).getText_2(1).OleValue.toString().trim())
        var ToReconcile = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
        Sys.HighlightObject(ToReconcile);
        if(ToReconcile.getSelection()){
          TextUtils.writeLog("To Reconcile is already is selected for Trancation No: "+ Transaction_No)
          
   var Date = table.getItem(v).getText_2(0).OleValue.toString().trim();
   var Trans_No = table.getItem(v).getText_2(1).OleValue.toString().trim();
   var Journal = table.getItem(v).getText_2(10).OleValue.toString().trim();
   var TEXT = table.getItem(v).getText_2(3).OleValue.toString().trim();
   var DEBIT = table.getItem(v).getText_2(4).OleValue.toString().trim();
   var CREDIT = table.getItem(v).getText_2(5).OleValue.toString().trim();
   
        }else{ 
        ToReconcile.Click();
        aqUtils.Delay(3000, Indicator.Text);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }   
   var Date = table.getItem(v).getText_2(0).OleValue.toString().trim();
   var Trans_No = table.getItem(v).getText_2(1).OleValue.toString().trim();
   var Journal = table.getItem(v).getText_2(10).OleValue.toString().trim();
   var TEXT = table.getItem(v).getText_2(3).OleValue.toString().trim();
   var DEBIT = table.getItem(v).getText_2(4).OleValue.toString().trim();
   var CREDIT = table.getItem(v).getText_2(5).OleValue.toString().trim();
   
        }
        
        if(v<table.getItemCount()-2)
          table.Keys("[Down]");

        }
        else{ 
        //  UnMarking To-Reconcile Check Box for other then selected Transaction No
          if(Loop){
        var Date = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
        Sys.HighlightObject(Date);
        Date.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
        aqUtils.Delay(5000, Indicator.Text);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
        Loop = false;
        }
        
        
        var ToReconcile = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPlainCheckboxView.Button;
        Sys.HighlightObject(ToReconcile);
        if(ToReconcile.getSelection()){
        ToReconcile.Click();
        aqUtils.Delay(3000, Indicator.Text);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }   
        }
        
        Log.Message(v)
        Log.Message(table.getItemCount()-1)
        if(v<table.getItemCount()-2)
          table.Keys("[Down]");
          
        }
      }

   if(flag){ 
   aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
   var Save = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
   Sys.HighlightObject(Save);
   Save.Click();
   aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);
   
   
   var AccNo = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget.getText().OleValue.toString().trim();
   var Local_Account_No = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget.getText().OleValue.toString().trim();

   
   
var ReconcilationSheet = "Banking Reconciliation MPL";  
   
ExcelUtils.setExcelName(workBook,ReconcilationSheet, true);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Bank Acc. No",ReconcilationSheet,Account_No);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Account",ReconcilationSheet,AccNo);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Local Account No",ReconcilationSheet,Local_Account_No);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Reconcilition Date",ReconcilationSheet,Date);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Trans No",ReconcilationSheet,Trans_No);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Journal",ReconcilationSheet,Journal);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"TEXT",ReconcilationSheet,TEXT);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"DEBIT",ReconcilationSheet,DEBIT);
ExcelUtils.WriteExcelSheet(EnvParams.Opco,"CREDIT",ReconcilationSheet,CREDIT);
//ExcelUtils.WriteExcelSheet(EnvParams.Opco,"BALANCE",ReconcilationSheet,BALANCE);
   
   

var Close_Balance_Calculated = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.McTextWidget;
Close_Balance_Calculated.Click();
Close_Balance_Calculated = Close_Balance_Calculated.getText();

var stmnt_Date = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.McDatePickerWidget;
var stmnt_No = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite2.McTextWidget;
var Close_Balance = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite3.McTextWidget;

stmnt_Date.Click();
//stmnt_Date.Keys(" ");
stmnt_Date.setText(Statment_Date);
aqUtils.Delay(2000, Indicator.Text);
stmnt_No.Click();
//stmnt_No.Keys(" ");
stmnt_No.setText(Statment_No);
aqUtils.Delay(2000, Indicator.Text);
Close_Balance.Click();
//Close_Balance.Keys(" ");
Close_Balance.setText(Close_Balance_Calculated);
aqUtils.Delay(2000, Indicator.Text);

var Save = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
Sys.HighlightObject(Save);
Save.Click();
   aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);
   
   
   var print = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
   Sys.HighlightObject(print);
   print.Click();
   aqUtils.Delay(5000, Indicator.Text);
   
   var SaveTitle = "";
var sFolder = "";

var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_BankReconciliation"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_BankReconciliation"+"*", 1).WndCaption.indexOf("P_BankReconciliation")!=-1){
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
Sys.HighlightObject(pdf);

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

filepathforMplValidation =sFolder+SaveTitle+".pdf";
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
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
ValidationUtils.verify(true,true,"Print Bank_Reconciliation In-Progress PDF is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Bank_Reconciliation In-Progress PDF",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")

var Approve_Reconciliation = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl2;
Sys.HighlightObject(Approve_Reconciliation);
Approve_Reconciliation.Click();
   aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);
    
   }
   
   
  }
  
  
}



function Reprint_Previous_Reconciliations(){ 
  
var History = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(History);
History.Click();
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);

   
var Bank_Acc = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(Bank_Acc);
Bank_Acc.Click();

   Bank_Acc.Keys(Account_No);
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   Bank_Acc.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]")
var Company = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(Company);
Company.Click();

   Company.Keys(EnvParams.Opco);
   aqUtils.Delay(2000, Indicator.Text);
   Company.Keys("[Tab]");
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Stat_Date = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Stat_Date);
Stat_Date.Click();

   Stat_Date.Keys(Statment_Date);
   aqUtils.Delay(2000, Indicator.Text);
   Stat_Date.Keys("[Tab]");
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Stat_No = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
Sys.HighlightObject(Stat_No);
Stat_No.Click();

   Stat_No.Keys(Statment_No);
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(2000, Indicator.Text);
   
   var table = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table)
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 

    if((table.getItem(v).getText_2(0).OleValue.toString().trim()==Account_No) && (table.getItem(v).getText_2(6).OleValue.toString().trim()==EnvParams.Opco) 
    && (table.getItem(v).getText_2(8).OleValue.toString().trim()==Statment_No)){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }  
  
  
  if(flag){
var CloseFilter = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(CloseFilter);
CloseFilter.Click();
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(2000, Indicator.Text);

var Re_Print_Reconciliation = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl3;
Sys.HighlightObject(Re_Print_Reconciliation);
Re_Print_Reconciliation.Click();


   aqUtils.Delay(5000, Indicator.Text);
   
   var SaveTitle = "";
var sFolder = "";

var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_BankReconciliation"+"*"+".pdf - Adobe Acrobat Reader DC*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_BankReconciliation"+"*"+".pdf - Adobe Acrobat Reader DC*", 1).WndCaption.indexOf("P_BankReconciliation")!=-1){
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
Sys.HighlightObject(pdf);

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

filepathforMplValidation =sFolder+SaveTitle+".pdf";
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
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
ValidationUtils.verify(true,true,"RePrint Bank_Reconciliation  PDF is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Bank_Reconciliation PDF",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")
   aqUtils.Delay(5000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);
}

   
}