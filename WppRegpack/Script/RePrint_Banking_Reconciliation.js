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
  if(Statment_Date == "AUTOFILL")
    getSpecificDate(0);
  if((Statment_Date=="") && (Statment_Date=="") ){
    ValidationUtils.verify(true,false,"Statement Date is needed to create Banking Reconciliation")
  }
  
  Statment_No = ExcelUtils.getRowDatas("Statement No",EnvParams.Opco);
  if((Statment_No=="") && (Statment_No=="") ){
    ValidationUtils.verify(true,false,"Statement Number is needed to create Banking Reconciliation")
  }
    
}


function Reprint_Previous_Reconciliations(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
     
var History = 
//Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy*").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
Sys.HighlightObject(History);
History.Click();
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   aqUtils.Delay(5000, Indicator.Text);
Sys.Desktop.Keys("[Up]");
var historyGrid = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(historyGrid); 
     
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

var Stat_Date = Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "", 4)
Sys.HighlightObject(Stat_Date);
Stat_Date.Click();

   //Stat_Date.Keys(Statment_Date);
   aqUtils.Delay(2000, Indicator.Text);
   Stat_Date.Keys("[Tab]");
   aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

//var Stat_No = 
////Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
//Aliases.Maconomy.Banking_Reconciliation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "", 4)
//
//Sys.HighlightObject(Stat_No);
//Stat_No.Click();
//
//   Stat_No.Keys(Statment_No);
//   aqUtils.Delay(2000, Indicator.Text);
//   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
//   aqUtils.Delay(2000, Indicator.Text);
   
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
   
  WorkspaceUtils.savePDF_And_WriteToExcel("RePrint_Bank_Reconciliation PDF","P_BankReconciliation");
}

   
}