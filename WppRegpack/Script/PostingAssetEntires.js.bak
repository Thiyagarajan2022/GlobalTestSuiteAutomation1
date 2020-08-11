//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "FixedAssetVal";
var AssetNo,Transactiontypeaddr,Amountbase;
var company,transactiontype = "";
var Language = "";
var level =0;
  Indicator.Show();
  Indicator.PushText("waiting for window to open");

 
function fixedassestpost(){
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
      sheetName = "FixedAssetVal";
      ExcelUtils.setExcelName(workBook, sheetName, true); 
      
          // AssetNo="";
  var workBook = "C:\\WppRegression_v12.50\\WppRegression_v12.50\\WppRegPack\\Testing Type\\SysTest\\DS_CHN_SYSTEST.xlsx"
var sheetName = "FixedAssetVal";
ExcelUtils.setExcelName(workBook, sheetName, true);
  
//AssetNo = ExcelUtils.getRowDatas("AssetNo",EnvParams.Opco)

//AssetNo="130710202";
//  if((AssetNo=="")||(AssetNo==null)){
////  jobNumber = readlog();
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  AssetNo = ExcelUtils.ReadExcelSheet("AssetNo",EnvParams.Opco,"Data Management");
//  Log.Message(AssetNo);
//  }
    goToJobMenuItem();   
     fixedassetvaladdr();
    getDetails();  
   fixedassetdrevlinfo();
   goToasset();
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
    
    var registrations=Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(registrations);
    registrations.Click();
    Delay(1000);
    
    
    delay(500);
var adjustment=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(adjustment);
adjustment.Click();

var newadj=Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(newadj);
newadj.Click();
    
    }
    
    function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
company = ExcelUtils.getRowDatas("CompanyNo",EnvParams.Opco)
Log.Message(company);
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"company is Needed to Create a fixedasset");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
transactiontype = ExcelUtils.getRowDatas("TransactionType",EnvParams.Opco)
Log.Message(company);
if((transactiontype==null)||(transactiontype=="")){ 
ValidationUtils.verify(false,true,"TransactionType is Needed to Create a fixedasset");


}

ExcelUtils.setExcelName(workBook, sheetName, true);
AssetNo = ExcelUtils.getRowDatas("AssetNo",EnvParams.Opco)
Log.Message(company);
if((AssetNo==null)||(AssetNo=="")){ 
ValidationUtils.verify(false,true,"AssetNo is Needed to Create a fixedasset");


}

ExcelUtils.setExcelName(workBook, sheetName, true);
Transactiontypeaddr = ExcelUtils.getRowDatas("AssetTransactionType",EnvParams.Opco)
Log.Message(company);
if((Transactiontypeaddr==null)||(Transactiontypeaddr=="")){ 
ValidationUtils.verify(false,true,"AssetTransactionType is Needed to Create a fixedasset");


}

ExcelUtils.setExcelName(workBook, sheetName, true);
Amountbase = ExcelUtils.getRowDatas("Amount",EnvParams.Opco)
Log.Message(company);
if((Amountbase==null)||(Amountbase=="")){ 
ValidationUtils.verify(false,true,"Amount is Needed to Create a fixedasset");


}

}


function fixedassetvaladdr(){
  
Delay(4000);
Sys.Process("Maconomy").Refresh();

var company1= Aliases.Maconomy.Screen6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget.getText();
Sys.HighlightObject(company1);

if(company1!="Company")
ValidationUtils.verify(false,true,"Company field is missing in Maconomy");
else
ValidationUtils.verify(true,true,"Company field is available in Maconomy");

var transactiontype1= Aliases.Maconomy.Screen6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget.getText();
Sys.HighlightObject(transactiontype1);
if(transactiontype1!="Transaction Type")
ValidationUtils.verify(false,true,"Transaction Type field is missing in Maconomy");
else
ValidationUtils.verify(true,true,"Transaction Type field is available in Maconomy");

}
    

function fixedassetdrevlinfo()
{

 var companyaddr = Aliases.Maconomy.Screen6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget2;
   if(company!=""){
        companyaddr.Click();                
        WorkspaceUtils.SearchByValueTableComp(companyaddr,"Company",company,"Company");
      }
   else{
        ValidationUtils.verify(true,true,"Company Number is Needed to Create Asset Adjustment");
      }  
      
   var transactionaddr = Aliases.Maconomy.Screen6.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
   //Sys.HighlightObject(transaction);   
   if(transactiontype!=""){
     transactionaddr.Click();  
        WorkspaceUtils.SearchByValueasset(transactionaddr,"Transaction Type",transactiontype,"Transactiontype");
     }
     else{ 
        ValidationUtils.verify(false,true,"Transaction Type is Needed to Create a Asset Adjustment");
    }   
      
    
    var createbtn = Aliases.Maconomy.Screen6.Composite.Composite.Composite2.Composite.Button;
//Sys.HighlightObject(createbtn);     
//reatebtn.Click();
  
 Sys.HighlightObject(createbtn);
  if(createbtn.isEnabled()){   
  createbtn.HoverMouse();
  ReportUtils.logStep_Screenshot(""); 
    createbtn.Click();
    ValidationUtils.verify(true,true,"Asset Adjustment is CREATED");   
  } 
  else{
    var cancelbtn = Aliases.Maconomy.Screen4.Composite.Composite.Composite2.Composite.cancel;
    Sys.HighlightObject(cancelbtn)    
    cancelbtn.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    cancelbtn.Click();
    ValidationUtils.verify(true,false,"Asset Adjustment is not Created");
    ReportUtils.logStep("ERROR","Asset Adjustment is not Created");
  } 
  aqUtils.Delay(4000, Indicator.Text);
  //closefilter
  Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46); 
                    
          
}

function goToasset()
{
var entries = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(entries);
var addbutton = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  addbutton.Click(); 

  var firstcell = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
  firstcell.Click();
  firstcell.Keys("[Tab][Tab][Tab]");
  aqUtils.Delay(1000,Indicator.Text);
Sys.Process("Maconomy").Refresh()  
  var assetno = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  if(AssetNo!=""){  
    assetno.Click();
    WorkspaceUtils.SearchByValueTableComp(assetno,"Asset",AssetNo,"AssetNo");    
  }
  else{
        ValidationUtils.verify(true,true,"Company Number is Needed to Create Asset Adjustment");
      }
      
  assetno.Keys("[Tab]");
  aqUtils.Delay(1000,Indicator.Text);
  Sys.Process("Maconomy").Refresh()
   var assettype = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
  assettype.Keys(" ");
  if(Transactiontypeaddr!=""){
    assettype.Click();aqUtils.Delay(1000, Indicator.Text);
       WorkspaceUtils.DropDownList(Transactiontypeaddr,"Asset Transaction Type");
       aqUtils.Delay(1000, Indicator.Text); 
    } 
    else{
      ValidationUtils.verify(false,true,"Asset Transaction Type is Needed to Create a Expense Sheet");  
    } 
    assettype.Keys("[Tab]");
  aqUtils.Delay(1000,Indicator.Text);
  
  var amountbaseaddr= Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
  if(Amountbase!=""){  
    amountbaseaddr.Click();
     //amount.Click();
    amountbaseaddr.setText(Amountbase);
    aqUtils.Delay(1000, Indicator.Text);
  
      save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
      Sys.HighlightObject(save); 
      save.Click();
      
approve = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
Sys.HighlightObject(approve);     
approve.Click();
 aqUtils.Delay(1000, Indicator.Text);
 
    var home = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
    home.Click();
    aqUtils.Delay(5000, Indicator.Text);
    
    var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    Sys.HighlightObject(table);
    var company2 = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
    company2.Click();
    company2.setText(company);
    company2.Keys("[Tab]");
    aqUtils.Delay(5000, Indicator.Text);
    
    var asset = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
    asset.Click();
    asset.setText(AssetNo); 
     var flag =false;
    for(var i=0;i<table.getItemCount();i++){
    if(table.getItem(i).getText_2(1).OleValue.toString().trim()==AssetNo){
      flag = true;
      ReportUtils.logStep_Screenshot("");
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
         Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
      break;
    } 
    else{
      table.Keys("[Down]");
    } 
  }  
  var closefilter= Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  ReportUtils.logStep_Screenshot("");
  ValidationUtils.verify(true,true,"Created Asset Adjustment is available in system");  
  ReportUtils.logStep("INFO", "Created Expenses is listed in table");
  
  
} 
}
