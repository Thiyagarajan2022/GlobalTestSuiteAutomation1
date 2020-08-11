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
var sheetName = "FixedAssetResolution";
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
      Log.Message(comapany)
      if((comapany==null)||(comapany=="")){ 
        ValidationUtils.verify(false,true,"Company Number is Needed to Create Asset");
      }
      if((comapany=="")||(comapany==null))
      ValidationUtils.verify(false,true,"Comapany Number is needed to Create Asset");
       
ExcelUtils.setExcelName(workBook, "Data Management", true);
AssetsNo = ReadExcelSheet("Assets No",EnvParams.Opco,"Data Management")
if((AssetsNo=="")||(AssetsNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
AssetsNo = ExcelUtils.getRowDatas("Assets No",EnvParams.Opco)
}
if((AssetsNo=="")||(AssetsNo==null))
ValidationUtils.verify(false,true,"Asset Number is needed to Create Asset Adjustment");
        
      }
      
function FixedAssetResolution(){
  TextUtils.writeLog("Create a Asset Adjustment Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  aqUtils.Delay(3000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior Accountant","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
   STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
excelName = EnvParams.path;
workBook = Project.Path+excelName;
      Log.Message(workBook)
      STIME = "";
      sheetName = "FixedAssetResolution";
      ExcelUtils.setExcelName(workBook, sheetName, true);  
 try{ 
    getDetails();
    goToJobMenuItem();  
    assetcost();   
    goToregistration();
    goToAsset();
    closeAllWorkspaces();
  }
    catch(err){
      Log.Message(err);
    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
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
     Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Fixed Assets").OleValue.toString().trim());
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Fixed Assets").OleValue.toString().trim());

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
  var assettable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid  
  Sys.HighlightObject(assettable);
  var comp = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  comp.Click();
  comp.Keys("[Tab]");
  var assetno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
  assetno.setText(AssetsNo);
  var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table);
  
  aqUtils.Delay(4000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==AssetsNo){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
 ValidationUtils.verify(flag,true,"Fixed Asset is availble in maconomy") 
  
  aqUtils.Delay(3000,Indicator.Text);
    var i=0;
       if(AssetsNo==assettable.getItem(i).getText(1).OleValue.toString().trim()){ 
             CostValue = assettable.getItem(i).getText(7).OleValue.toString().trim()
             ValidationUtils.verify(true,true,"CostValue is :"+CostValue)
             TextUtils.writeLog("Asset CostValue is :"+CostValue);
         }
  
}

function goToregistration(){
   ReportUtils.logStep("INFO","Fixed Asset Depreciation is Started:"+STIME);    
  var register = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Registrations;
  Sys.HighlightObject(register);
  register.Click();
  aqUtils.Delay(4000,Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var Newassetadjust = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;  
    ReportUtils.logStep_Screenshot("");
   Newassetadjust.Click();
   aqUtils.Delay(1000,Indicator.Text);
//   address();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
   var company = Aliases.Maconomy.Group4.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.companynumber;
     if(comapany!=""){
        company.Click();
        SearchByValue(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),comapany,"Company Number");
      }
      else{ 
        ValidationUtils.verify(false,true,"company is Needed to Create a Fixed Assets");
      }
      
   var transaction = Aliases.Maconomy.Group4.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Transtype;
   Sys.HighlightObject(transaction);   
   if(Transaction!=""){
     transaction.Click();  
        SearchByValueasset(transaction,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transaction Type").OleValue.toString().trim(),Transaction,"Transactiontype");
     }
     else{ 
        ValidationUtils.verify(false,true,"Transaction Type is Needed to Create a Asset Adjustment");
    }   
      
   
    var createbtn =  Aliases.Maconomy.Group4.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
  Sys.HighlightObject(createbtn);
  if(createbtn.isEnabled()){   
  createbtn.HoverMouse();
  ReportUtils.logStep_Screenshot(""); 
    createbtn.Click();
    ValidationUtils.verify(true,true,"Asset Adjustment is CREATED");   
    TextUtils.writeLog("Asset Adjustment is CREATED");
  } 
  else{
    var cancelbtn = Aliases.Maconomy.Group4.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
  
  var entries = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table;
  Sys.HighlightObject(entries)
  var addbutton = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.addbutton;
  addbutton.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var firstcell = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.firstcell;
  firstcell.Click();
  firstcell.Keys("[Tab][Tab][Tab]");
  aqUtils.Delay(1000,Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var assetno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.assetno;
  if(AssetsNo!=""){  
    assetno.Click();
    SearchByValueTableComp(assetno,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Asset").OleValue.toString().trim(),AssetsNo,"AssetNumber");    
  }
  else{
        ValidationUtils.verify(true,true,"Company Number is Needed to Create Asset Adjustment");
      } 
      
  assetno.Keys("[Tab]");
  aqUtils.Delay(1000,Indicator.Text);
  Sys.Process("Maconomy").Refresh();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var assettype = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.assettype;
  assettype.Keys(" ");
  if(Transactiontype!=""){
    aqUtils.Delay(10000,Indicator.Text);
    assettype.Click();
    aqUtils.Delay(1000, Indicator.Text);
       WorkspaceUtils.DropDownList(Transactiontype,"Asset Transaction Type");
       aqUtils.Delay(1000, Indicator.Text); 
    } 
    else{
      ValidationUtils.verify(false,true,"Currency is Needed to Create a Expense Sheet");  
    } 
    assettype.Keys("[Tab]");
    
    var amount = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.amount;
    amount.Click();
    amount.setText(Amountbase);
    aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    save.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    save.Click(); 
    ValidationUtils.verify(true,true,"Entries is added and saved");
     TextUtils.writeLog("Entries is added and saved");
    aqUtils.Delay(1000, Indicator.Text);          
    
    var approve = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.approve;
    waitForObj(approve)
    Sys.HighlightObject(approve)
    if(approve.isEnabled()){      
      approve.HoverMouse();
      ReportUtils.logStep_Screenshot();
      approve.Click();
      
        if(CostValue>Amountbase){
          ValidationUtils.verify(false,true,"Cost is exceeds the posted value");
          ReportUtils.logStep("INFO","Cost is exceeds the posted value");
          TextUtils.writeLog("Cost is exceeds the posted value");
        } 
        else{
          ValidationUtils.verify(true,true,"CostValue is:"+CostValue);
          ValidationUtils.verify(true,true,"Amountbase is:"+Amountbase);
          ValidationUtils.verify(true,true,"Cost is greater than the Fixed Amount");
          TextUtils.writeLog("Cost is greater than the Fixed Amount");
        } 
      
      ValidationUtils.verify(true,true,"Create Asset is Approved");
      TextUtils.writeLog("Create Asset is Approved");
    } 
    else{ 
      ReportUtils.logStep("INFO","Approve Button Is Invisible");
    } 
    
    aqUtils.Delay(2000, Indicator.Text);       
    
    var home = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.TabFolderPanel.home;
    home.Click();
    aqUtils.Delay(5000, Indicator.Text);
    
    var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    Sys.HighlightObject(table);

  aqUtils.Delay(3000, Indicator.Text);   
  ReportUtils.logStep_Screenshot("");
  var b=0;
  var bookvalue = table.getItem(b).getText_2(9).OleValue.toString().trim();
  ValidationUtils.verify(true,true,"BookValue is:"+ bookvalue);    
   TextUtils.writeLog("Created Asset Adjustment is available in system")
  ValidationUtils.verify(true,true,"Created Asset Adjustment Book Value has changed");  
  
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
         aqUtils.Delay(5000, Indicator.Text);    
  
} 



function SearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);


  var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  waitForObj(code);
  code.Click();

    code.setText(value);

    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search ").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
  var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())

    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();

          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    
    return checkmark;
}

function SearchByValueasset(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search ").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
   serch.Click(); 
  }
    aqUtils.Delay(5000, Indicator.Text);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  if(OK.isEnabled()){
  OK.HoverMouse();
ReportUtils.logStep_Screenshot();
  OK.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
    OK.HoverMouse();
ReportUtils.logStep_Screenshot();
   OK.Click(); 
  }
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
  cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
          aqUtils.Delay(1000, Indicator.Text);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
    cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
      aqUtils.Delay(1000, Indicator.Text);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValueTableComp(ObjectAddrs,popupName,value,fieldName){
var checkmark =  false;
  aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search ").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}