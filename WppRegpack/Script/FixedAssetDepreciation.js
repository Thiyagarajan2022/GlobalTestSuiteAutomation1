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
var sheetName = "FixedAssetDepreciation";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");

var comapany= "";
var AssetsNo="";
var Transaction="";
var Transactiontype="";
 var Amountbase="";
 var Assetcredate = "";
 var CostValue = "";
 var STIME = "";
  
function getDetails(){
    ExcelUtils.setExcelName(workBook, sheetName, true);
      AssetsNo = ExcelUtils.getRowDatas("Assets No",EnvParams.Opco)
        if((AssetsNo=="")||(AssetsNo==null)){
              ExcelUtils.setExcelName(workBook, "Data Management", true);
              AssetsNo = ReadExcelSheet("Assets No",EnvParams.Opco,"Data Management")
        } 
        if((AssetsNo=="")||(AssetsNo==null))
        ValidationUtils.verify(false,true,"Asset Number is needed to Create Asset Adjustment");
        
//        CostValue = ExcelUtils.getRowDatas("Cost",EnvParams.Opco)
//        Log.Message(CostValue)
//        if((CostValue=="")||(CostValue==null)){
//          ExcelUtils.setExcelName(workBook, "Data Management", true);
//          CostValue = ReadExcelSheet("Cost",EnvParams.Opco,"Data Management");
//          Log.Message(CostValue)
//        }      
                 
        Transaction = ExcelUtils.getRowDatas("Transaction Type",EnvParams.Opco)
        if((Transaction==null)||(Transaction=="")){ 
        ValidationUtils.verify(false,true,"Transaction Type is Needed to Create a Asset Adjustment");
        } 
        Transactiontype = ExcelUtils.getRowDatas("Asset Transaction type",EnvParams.Opco)
        if((Transactiontype==null)||(Transactiontype=="")){ 
        ValidationUtils.verify(false,true,"Asset Transaction type is Needed to Create a Asset Adjustment");
        } 
        Amountbase = ExcelUtils.getRowDatas("Amount",EnvParams.Opco)
        if((Amountbase==null)||(Amountbase=="")){ 
        ValidationUtils.verify(false,true,"Amount is Needed to Create a Asset Adjustment");
        }   
        comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
        if((comapany==null)||(comapany=="")){ 
        ValidationUtils.verify(false,true,"company is Needed to Create a Asset Adjustment");
        }  
      }
      
function fixedassest(){
      Language = "";
      Language = EnvParams.Language;
        if((Language==null)||(Language=="")){
          ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
        }      
      Language = EnvParams.LanChange(Language);
      WorkspaceUtils.Language = Language;
      Log.Message(Language)
      STIME = WorkspaceUtils.StartTime();
      excelName = EnvParams.path;
      workBook = Project.Path+excelName;
      sheetName = "FixedAssetDepreciation";
      ExcelUtils.setExcelName(workBook, sheetName, true);  
    getDetails();
    goToJobMenuItem();  
    assetcost();   
    goToregistration();
    goToAsset();
    closeAllWorkspaces();
} 


function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet0.Assets.Exists()){
          ImageRepository.ImageSet0.Assets.Click();// GL
          }
          else if(ImageRepository.ImageSet0.Assets1.Exists()){
          ImageRepository.ImageSet0.Assets1.Click();
          }
          else{
          ImageRepository.ImageSet0.Assets2.Click();
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
    aqUtils.Delay(3000,Indicator.Text);
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
    aqUtils.Delay(3000,Indicator.Text);
  }
  
  
function address(){
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Process("Maconomy").Refresh();
    var companylable = Aliases.Maconomy.Group4.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.lablecompany.getText().OleValue.toString().trim();
    if(companylable!="Company")
    ValidationUtils.verify(false,true,"Company field is missing in macanomy for the Create Asset Adjustment");
    else
    ValidationUtils.verify(true,true,"Company field is available in Macanomy for the Create Asset Adjustment");

    var TransactionType = Aliases.Maconomy.Group4.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.transtypee.getText().OleValue.toString().trim();
    if(TransactionType!="Transaction Type")
    ValidationUtils.verify(false,true,"Transaction Type field is missing in macanomy for the Create Asset Adjustment");
    else
    ValidationUtils.verify(true,true,"Transaction Type field is available in Macanomy for the Create Asset Adjustment");
}

function assetcost(){
  var assettable = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(assettable);
  var comp = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  comp.Click();
  comp.Keys("[Tab]");
  var assetno = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
  assetno.setText(AssetsNo);
  aqUtils.Delay(3000,Indicator.Text);
    var i=0;
       if(AssetsNo==assettable.getItem(i).getText(1).OleValue.toString().trim()){ 
             CostValue = assettable.getItem(i).getText(7).OleValue.toString().trim()
             ValidationUtils.verify(true,true,"CostValue is :"+CostValue)
         }
  
}

function goToregistration(){
   ReportUtils.logStep("INFO","Fixed Asset Depreciation is Started:"+STIME);    
  var register = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Registrations;
  Sys.HighlightObject(register);
  register.Click();
  aqUtils.Delay(4000,Indicator.Text);
  
  var Newassetadjust = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.newassetadjust;
//   Newassetadjust.HoverMouse();
    ReportUtils.logStep_Screenshot("");
   Newassetadjust.Click();
   aqUtils.Delay(1000,Indicator.Text);
   address();
   
   var company = Aliases.Maconomy.Group4.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.companynumber;
   if(comapany!=""){
        company.Click();                
//        WorkspaceUtils.SearchByValueTableComp(company,"Company",comapany,"Company");
          WorkspaceUtils.SearchByValue(company,"Company",comapany,"Company Number");
      }
   else{
        ValidationUtils.verify(true,true,"Company Number is Needed to Create Asset Adjustment");
      }  
      
   var transaction = Aliases.Maconomy.Group4.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Transtype;
   Sys.HighlightObject(transaction);   
   if(Transaction!=""){
     transaction.Click();  
        WorkspaceUtils.SearchByValueasset(transaction,"Transaction Type",Transaction,"Transactiontype");
     }
     else{ 
        ValidationUtils.verify(false,true,"Transaction Type is Needed to Create a Asset Adjustment");
    }   
      
   
    var createbtn =  Aliases.Maconomy.Group4.Composite.Composite.Composite2.Composite.create;
  Sys.HighlightObject(createbtn);
  if(createbtn.isEnabled()){   
  createbtn.HoverMouse();
  ReportUtils.logStep_Screenshot(""); 
    createbtn.Click();
    ValidationUtils.verify(true,true,"Asset Adjustment is CREATED");   
  } 
  else{
    var cancelbtn = Aliases.Maconomy.Group4.Composite.Composite.Composite2.Composite.cancel;
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

function goToAsset(){
  
  var entries = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table;
  Sys.HighlightObject(entries)
  var addbutton = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.addbutton;
  addbutton.Click();
  
  var firstcell = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.firstcell;
  firstcell.Click();
  firstcell.Keys("[Tab][Tab][Tab]");
  aqUtils.Delay(1000,Indicator.Text);
  
  var assetno = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.assetno;
  if(AssetsNo!=""){  
    assetno.Click();
    WorkspaceUtils.SearchByValueTableComp(assetno,"Asset",AssetsNo,"AssetNumber");    
  }
  else{
        ValidationUtils.verify(true,true,"Company Number is Needed to Create Asset Adjustment");
      } 
      
  assetno.Keys("[Tab]");
  aqUtils.Delay(1000,Indicator.Text);
//  Sys.Process("Maconomy").Refresh();
  var assettype = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.assettype;
  assettype.Keys(" ");
  if(Transactiontype!=""){
    assettype.Click();aqUtils.Delay(1000, Indicator.Text);
       WorkspaceUtils.DropDownList(Transactiontype,"Asset Transaction Type");
       aqUtils.Delay(1000, Indicator.Text); 
    } 
    else{
      ValidationUtils.verify(false,true,"Currency is Needed to Create a Expense Sheet");  
    } 
    assettype.Keys("[Tab]");
    
    var amount = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.amount;
    amount.Click();
    amount.setText(Amountbase);
    aqUtils.Delay(1000, Indicator.Text);
    
    var save = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    save.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    save.Click(); 
    ValidationUtils.verify(true,true,"Entries is added and saved");
    aqUtils.Delay(1000, Indicator.Text);          
    
    var approve = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.approve;
    if(approve.isEnabled()){      
      approve.HoverMouse();
      ReportUtils.logStep_Screenshot();
      approve.Click();
      
        if(CostValue>Amountbase){
          ValidationUtils.verify(true,true,"CostValue is:"+CostValue);
          ValidationUtils.verify(true,true,"Amountbase is:"+Amountbase);
          ValidationUtils.verify(true,true,"Cost is greater than the Fixed Amount");
        } 
        else{
          ValidationUtils.verify(false,true,"Cost is exceeds the posted value");
          ReportUtils.logStep("INFO","Cost is exceeds the posted value");
        } 
      
      ValidationUtils.verify(true,true,"Create Asset is Approved");
    } 
    else{ 
      ReportUtils.logStep("INFO","Approve Button Is Invisible");
    } 
    
    aqUtils.Delay(2000, Indicator.Text);       
    
    var home = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.TabFolderPanel.home;
    home.Click();
    aqUtils.Delay(5000, Indicator.Text);
    
    var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    Sys.HighlightObject(table);

    var company = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.company;
    company.Click();
    company.setText(comapany);
    company.Keys("[Tab]");
    aqUtils.Delay(5000, Indicator.Text);
    
    var asset = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.asset;
    asset.Click();
    asset.setText(AssetsNo); 
     var flag =false;
    for(var i=0;i<table.getItemCount();i++){
    if(table.getItem(i).getText_2(1).OleValue.toString().trim()==AssetsNo){
      flag = true;
      ReportUtils.logStep_Screenshot("");          
    } 
    else{
      table.Keys("[Down]");
    } 
  } 
  aqUtils.Delay(3000, Indicator.Text);   
  ReportUtils.logStep_Screenshot("");
  var b=0;
  var bookvalue = table.getItem(b).getText_2(9).OleValue.toString().trim();
  ValidationUtils.verify(true,true,"BookValue is:"+ bookvalue);    
  ValidationUtils.verify(true,true,"Created Asset Adjustment is available in system");  
  
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
         aqUtils.Delay(5000, Indicator.Text);    
  
} 


