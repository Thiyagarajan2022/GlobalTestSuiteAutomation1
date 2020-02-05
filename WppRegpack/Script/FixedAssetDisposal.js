﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

 
var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "FixedAssetDisposal";
var AssetNo;
var amountsold,percentagesold,datesale,remark = "";
var Language = "";
var level =0;
  Indicator.Show();
  Indicator.PushText("waiting for window to open");


function fixedassestdisposal(){
      Language = "";
      Language = EnvParams.Language;
        if((Language==null)||(Language=="")){
          ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
        }      
      Language = EnvParams.LanChange(Language);
      WorkspaceUtils.Language = Language;
      Log.Message(Language)
      
      excelName = EnvParams.path;
      workBook = Project.Path+excelName;
      STIME = "";      
      sheetName = "FixedAssetDisposal";
      ExcelUtils.setExcelName(workBook, sheetName, true); 
      
           AssetNo="";
  var workBook = "C:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\DS_CHN_SYSTEST.xlsx"
var sheetName = "FixedAssetDisposal";
ExcelUtils.setExcelName(workBook, sheetName, true);
  
//AssetNo = ExcelUtils.getRowDatas("AssetNo",EnvParams.Opco)

AssetNo="130710202";
  if((AssetNo=="")||(AssetNo==null)){
//  jobNumber = readlog();
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  AssetNo = ExcelUtils.ReadExcelSheet("AssetNo",EnvParams.Opco,"Data Management");
  Log.Message(AssetNo);
  }
    goToJobMenuItem();     
    FixedassetScreen_Address();
     getDetails();
    fixedassetdinfo();
    registration();
   TransactionNo();
    //gotoTimeExpenses();
    //closeAllWorkspaces();
} 


function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet.Assets.Exists()){
          ImageRepository.ImageSet.Assets.Click();// GL
          }
          else if(ImageRepository.ImageSet.Assets1.Exists()){
          ImageRepository.ImageSet.Assets1.Click();
          }
          else{
          ImageRepository.ImageSet.Assets2.Click();
          }
    aqUtils.Delay(3000, Indicator.Text);
    Sys.Desktop.KeyDown(0x12);
    Sys.Desktop.KeyDown(0x20);
    Sys.Desktop.KeyUp(0x12);
    Sys.Desktop.KeyUp(0x20);
    Sys.Desktop.KeyDown(0x58);
    Sys.Desktop.KeyUp(0x58);  
    aqUtils.Delay(1000, Indicator.Text);
    var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
    var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
    Delay(3000);
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
        Client_Managt.ClickItem("|Fixed Assets");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|Fixed Assets");
      }
    }
    Delay(3000);
    
    var registrations=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(registrations);
    registrations.Click();
    Delay(1000);
    var disposal= Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.TabControl2;
    //Aliases.Maconomy.Screen2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
    //Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl2;
    Sys.HighlightObject(disposal);
    disposal.Click();
    
    
   var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstCell = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
    //firstCell.setText(EnvParams.Opco);
//  firstCell.Keys("1707");
firstCell.Keys(AssetNo+"[Tab]");
Log.Message(AssetNo);
var closefilter = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//var job = Sys.Process("Maconomy").SWTObject("Screen", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//Log.Message(EmployeeNo);
////  job.setText("GAIL C COUTINHO")
//job.setText(EmployeeNo);

Delay(6000);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(AssetNo)){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    

ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Asset Created is available in system");
   
  
if(flag){ 
closefilter.Click();
Delay(5000);

var approvesale=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2
Sys.HighlightObject(approvesale);
    approvesale.Click();
    
  }
  
}

//amountsold,percentagesold,datesale,remark
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
amountsold = ExcelUtils.getRowDatas("AmountSold",EnvParams.Opco)
Log.Message(amountsold);
if((amountsold==null)||(amountsold=="")){ 
ValidationUtils.verify(false,true,"amountsold is Needed to Create a fixedasset");
}
percentagesold = ExcelUtils.getRowDatas("PercentageSold",EnvParams.Opco)
Log.Message(percentagesold);
if((percentagesold==null)||(percentagesold=="")){ 
ValidationUtils.verify(false,true,"percentagesold is Needed to Create a fixedasset");
}
datesale = ExcelUtils.getRowDatas("DateofSale",EnvParams.Opco)
if((datesale==null)||(datesale=="")){ 
ValidationUtils.verify(false,true,"datesale is Needed to Create a fixedasset");
}

remark = ExcelUtils.getRowDatas("Remark",EnvParams.Opco)
Log.Message(remark);
if((remark==null)||(remark=="")){ 
ValidationUtils.verify(false,true,"remark is Needed to Create a fixedasset");
}

}

function FixedassetScreen_Address(){ 
//Checking Labels in Job Create Wizard
Delay(4000);
Sys.Process("Maconomy").Refresh();

var Amount_Sold1 = Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McTextWidget.getText();
if(Amount_Sold1!="Amount Sold")
ValidationUtils.verify(false,true,"Amount Sold field is missing in Maconomy");
else
ValidationUtils.verify(true,true,"Amount Sold field is available in Maconomy");

var Percentage_Sold1 = Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McTextWidget.getText();


Log.Message(Percentage_Sold1);
if(Percentage_Sold1!="Percentage Sold")
ValidationUtils.verify(false,true,"Percentage Sold field is missing in Maconomy");
else
ValidationUtils.verify(true,true,"Percentage Sold field is available in Maconomy");

var Date_Sale1 = Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McTextWidget.getText();
Log.Message(Date_Sale1);
if(Date_Sale1!="Date of Sale")
ValidationUtils.verify(false,true,"Date of Sale field is missing in Maconomy");
else
ValidationUtils.verify(true,true,"Date of Sale field is available in Maconomy");

var Remark_1 = Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite4.McTextWidget.getText();
Log.Message(Remark_1);
if(Remark_1!="Remark")
ValidationUtils.verify(false,true,"Remark field is missing in Maconomy");
else
ValidationUtils.verify(true,true,"Remark field is available in Maconomy");

}




function fixedassetdinfo()
{
  
if((amountsold!="") && (amountsold!=null)){
var Amount_Sale=Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McTextWidget2;
  Log.Message(amountsold);
Amount_Sale.Click();
Amount_Sale.setText(amountsold);
ValidationUtils.verify(true,true,"AmountSale is entered in Maconomy");
}else{ 
  ValidationUtils.verify(false,true,"AmountSale is Needed to Create a FixedAsset");
}

  
if((percentagesold!="") && (percentagesold!=null)){
var Percentage_Sold= Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McTextWidget2;
Log.Message(percentagesold);  
Percentage_Sold.Click();
Percentage_Sold.setText(percentagesold);
ValidationUtils.verify(true,true,"PercentageSold is entered in Maconomy");
}else{ 
  ValidationUtils.verify(false,true,"PercentageSold is Needed to Create a FixedAsset");
}

var datesale1 = Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite3.McDatePickerWidget;
if((datesale!="") && (datesale!=null)){
WorkspaceUtils.CalenderDateSelection(datesale1,datesale);
}
else{ 
  ValidationUtils.verify(false,true,"date of sale is Needed to Fixed Asset Disposal");
}

if((remark!="") && (remark!=null)){
var Remark1= Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite4.McTextWidget2;
//Aliases.Maconomy.Screen5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite2.McTextWidget2;
  Remark1.Click();
Remark1.setText(remark);
ValidationUtils.verify(true,true,"Remark is entered in Maconomy");
}else{ 
  ValidationUtils.verify(false,true,"Remark is Needed to Create a FixedAsset");
}

 var submitapprovalsale=Aliases.Maconomy.Screen5.Composite.Composite.Composite2.Composite.Button;
Sys.HighlightObject(submitapprovalsale);
submitapprovalsale.Click();
Delay(1000);
  //WorkspaceUtils.closeAllWorkspaces();
}



function registration()
{
  delay(1000);
  var reg1=Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(reg1);
reg1.Click();

delay(500);
var adjustment=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(adjustment);
adjustment.Click();

var alladjustment=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
Sys.HighlightObject(alladjustment);
alladjustment.Click();

var table= Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);

var firstcell=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(firstcell);
firstcell.Click();
delay(1000);
firstcell.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
delay(1000);
var approvedon=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget3;
approvedon.Click();
Log.Message(datesale);
approvedon.setText(datesale);


var closefilter=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 closefilter.Click();
 

}


function TransactionNo(){
//var table= Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
    //table.getText_2(1).Olev;
//    var approvedon=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget3;
//approvedon.Click() 
//    
    var entries = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
    //Aliases.Maconomy.Screen2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
entries.Click();
Delay(1000);
var table= Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(table);
Log.Message(table.getItemCount());
 //table.getText_2(1).Olev;
//   var amountsold="20.00";
//   var  percentagesold="10.00";
//   var remark="FixedAsset";
var number =""
//amountsold,percentagesold,datesale,remark
for(var i=0;i<table.getItemCount();i++)
{ 
//var getitem=table.getItemCount();
//if (getitem>0){
  // var entries = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
    //Aliases.Maconomy.Screen2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
//entries.Click();
//Delay(1000);
var table= Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  if((table.getItem(i).getText_2(5).OleValue.toString().trim()== amountsold) &&
  (table.getItem(i).getText_2(11).OleValue.toString().trim()== percentagesold) &&
  (table.getItem(i).getText_2(12).OleValue.toString().trim()== remark))
  { 
    number = table.getItem(i).getText_2(1).OleValue.toString().trim(); 
    ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("AssetTransactionNo",EnvParams.Opco,"Data Management",number)
// ValidationUtils.verify(false,true,"Number is matched");
  }
  //}
  else{ 
    table.Keys("[Down]")
    var showfilter=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(showfilter);
    showfilter.Click();
    Delay(1000);  
    
    
// ValidationUtils.verify(true,true,"Number not matched");
  }
  
  var table1=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    
   // Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    table1.Keys("[Down]")
    Delay(1000);
var closefilter=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 Sys.HighlightObject(closefilter);
closefilter.Click();
}
Log.Message(number);


// ExcelUtils.setExcelName(workBook,"Data Management", true);
//      ExcelUtils.WriteExcelSheet("AssetTransactionNo",EnvParams.Opco,"Data Management",number)

//if (number!=""){
//  closefilter
//}
//  else {
//    
//  }
    

}


function gl(){
  

}

//  function fixedassestdisposal()
//  {
//    level=0;
//EmployeeNo="";
//  var workBook = "C:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\DS_CHN_SYSTEST.xlsx"
//var sheetName = "ChangeEmployee";
//ExcelUtils.setExcelName(workBook, sheetName, true);
//  EmployeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
//  if((EmployeeNo=="")||(EmployeeNo==null)){
////  jobNumber = readlog();
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  EmployeeNo = ExcelUtils.ReadExcelSheet("Employee No",EnvParams.Opco,"Data Management");
//  }
//  getDetails();
//  goToMenu();
//  employeeinfo();
//  }

