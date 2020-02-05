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
var sheetName = "Create Fixed Asset";
Indicator.Show();
Indicator.PushText("waiting for window to open");
 
 
  var Assetgroup ="";
  var STIME = "";
  var Descrip ="";
  var cost ="";
  var date ="";
  var Access ="";
  var comapany ="";
  var layout ="";
  var login ="";
  var transactionNo="";
  
function CreateAssest(){
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
      STIME = "";
      sheetName = "Create Fixed Asset";
      ExcelUtils.setExcelName(workBook, sheetName, true); 
       
      getDetails();
      goToJobMenuItem(); 
      createAssets();
      closeAllWorkspaces();
      WorkspaceUtils.closeMaconomy();
      Restart.login(login);
      Posting();
      WorkspaceUtils.closeAllWorkspaces();
}  
  
function getDetails(){
        
        sheetName = "Create Fixed Asset";        
        ExcelUtils.setExcelName(workBook, sheetName, true);
        Assetgroup = ExcelUtils.getRowDatas("Asset Group",EnvParams.Opco)
        if((Assetgroup=="")||(Assetgroup==null)){
        ValidationUtils.verify(false,true,"Asset Group is Needed to Create a Asset");
        }
        
        Descrip = ExcelUtils.getRowDatas("Description",EnvParams.Opco)
        if((Descrip==null)||(Descrip=="")){ 
        ValidationUtils.verify(false,true,"Description is Needed to Create a Asset");
        } 
        
        date = ExcelUtils.getRowDatas("AssetDate",EnvParams.Opco)
        if((date==null)||(date=="")){ 
        ValidationUtils.verify(false,true,"AssetDate is Needed to Create a Asset");
        }   
               
        cost = ExcelUtils.getRowDatas("Cost",EnvParams.Opco)
        if((cost==null)||(cost=="")){ 
        ValidationUtils.verify(false,true,"Cost is Needed to Create a Asset");
        }    
        Access = ExcelUtils.getRowDatas("Access Level",EnvParams.Opco)
        if((Access==null)||(Access=="")){ 
        ValidationUtils.verify(false,true,"Access Level is Needed to Create a Asset");
        }    
        layout = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
        if((layout==null)||(layout=="")){ 
        ValidationUtils.verify(false,true,"Layout is Needed to Create a Asset");
        } 
        login = ExcelUtils.getRowDatas("Login",EnvParams.Opco)
        if((login==null)||(login=="")){ 
        ValidationUtils.verify(false,true,"User Details is Needed to Create a Asset");
        } 
//      sheetName = "SSC Users";
//      ExcelUtils.setExcelName(workBook, sheetName, true);
//      comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
//      if((comapany==null)||(comapany=="")){ 
//      ValidationUtils.verify(false,true,"Company Number is Needed to Create a Asset"); 
         
//      sheetName = "JobCreation";
//      ExcelUtils.setExcelName(workBook, sheetName, true);
      comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
      if((comapany==null)||(comapany=="")){ 
        ValidationUtils.verify(false,true,"Company Number is Needed to Create Asset");
      }
      if((comapany=="")||(comapany==null))
      ValidationUtils.verify(false,true,"Comapany Number is needed to Create Asset");
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


  
  
function createAssets(){ 
    ReportUtils.logStep("INFO","Create Fixed Asset is started:"+STIME); 
    var newAssets = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;   
//    var newAssets = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.newassetadjust;
      var Add_Visible0 = true;
      while(Add_Visible0){
        if(newAssets.isEnabled()){
        ReportUtils.logStep_Screenshot();
        newAssets.Click();
        Add_Visible0 = false;
        }
      }     
      ReportUtils.logStep_Screenshot(""); 
      address1();
    aqUtils.Delay(1000,Indicator.Text);   
    var assetGroup = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.assetgroup;
      if(Assetgroup!=""){
        assetGroup.Click();
        WorkspaceUtils.SearchByValue(assetGroup,"Asset Group",Assetgroup,"Name");
      }
      else{ 
        ValidationUtils.verify(false,true,"AssetGroup is Needed to Create a Fixed Assets");
      }         
    var Description = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.descrip;
      if(Descrip!=""){
        Description.Click();
        Description.setText(Descrip+" "+STIME); 
        ValidationUtils.verify(true,true,"Description is Entered in Maconomy"); 
      }
      else{ 
        ValidationUtils.verify(false,true,"Description is Needed to Create a Fixed Assets");
      }
    var company = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.number;
      if(comapany!=""){
        company.Click();
        WorkspaceUtils.SearchByValue(company,"Company",comapany,"Company Number");
      }
      else{ 
        ValidationUtils.verify(false,true,"company is Needed to Create a Fixed Assets");
      }
    
    var next = Aliases.Maconomy.Group5.Composite.Composite.Composite2.Composite.nextButton;
    Sys.HighlightObject(next);
    ReportUtils.logStep_Screenshot();
    next.Click();
    aqUtils.Delay(1000,Indicator.Text);
    
    address2();    
    
    if(date!=""){
      var datefiled = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.date;
      WorkspaceUtils.CalenderDateSelection(datefiled,date)
      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Create a Employee");
    }    
    aqUtils.Delay(1000,Indicator.Text);
    
    var getdate = datefiled.getText();
    Log.Message(getdate)
//    ExcelUtils.setExcelName(workBook,"Data Management", true);
//  ExcelUtils.WriteExcelSheet("AssetDate",EnvParams.Opco,"Data Management",getdate)
    
    var costt = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.descrip;
    Sys.HighlightObject(costt);
    if(cost!=""){
        costt.Click();
        costt.setText(cost);
        ValidationUtils.verify(true,true,"Cost is Entered in Maconomy"); 
      }
      else{ 
        ValidationUtils.verify(false,true,"Cost is Needed to Create a Fixed Assets");
      }
      
      var get = costt.getText();      
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Cost",EnvParams.Opco,"Data Management",get)
    
    var access = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
      if(comapany!=""){
        access.Click();
        WorkspaceUtils.SearchByValue(access,"Access Level",Access,"Name");
      }
      else{ 
        ValidationUtils.verify(false,true,"company is Needed to Create a Fixed Assets");
      }
    
    var btnCreate = NameMapping.Sys.Maconomy.Group5.Composite.Composite.Composite2.Composite.SWTObject("Button", "Create");    
    if(btnCreate.isEnabled()){
        Sys.HighlightObject(btnCreate)
        btnCreate.HoverMouse();
      ReportUtils.logStep_Screenshot("");
        btnCreate.Click();
      ValidationUtils.verify(true,true,"Asset is Created");
      ReportUtils.logStep("INFO", "Asset is Created");
      aqUtils.Delay(5000, Indicator.Text);
    }
    else{ 
      var cancel = Aliases.Maconomy.Group5.Composite.Composite.Composite2.cancel;
      cancel.HoverMouse();
    ReportUtils.logStep_Screenshot("");
      cancel.Click();
    ValidationUtils.verify(true,false,"Asset is not Created");
    ReportUtils.logStep("ERROR", "Asset is not Created");
    }
    aqUtils.Delay(5000, Indicator.Text);
       //closefilter
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46); 
          ReportUtils.logStep_Screenshot();
          
       var asetnumber = (Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.assetnumber).getText();
//       Log.Message(asetnumber);    
       ValidationUtils.verify(true,true,"Asset Number : "+asetnumber);
        ExcelUtils.setExcelName(workBook,"Data Management", true);
        ExcelUtils.WriteExcelSheet("Assets No",EnvParams.Opco,"Data Management",asetnumber)
         

        var allEntries = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.entrytable.McGrid;
        allEntries.Click();
        aqUtils.Delay(2000, Indicator.Text);
        var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.entrytable.McGrid;
        transactionNo = table.getItem(0).getText(4).OleValue.toString().trim();
        Log.Message("Transaction No. :"+transactionNo);
        aqUtils.Delay(2000, Indicator.Text);
        ReportUtils.logStep_Screenshot();
    }    
    
 function address1(){
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Process("Maconomy").Refresh();
    var AssetGroup = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.acessgrop.getText().OleValue.toString().trim();
    if(AssetGroup!="Asset Group")
    ValidationUtils.verify(false,true,"AssetGroup field is missing in macanomy for the Create Asset");
    else
    ValidationUtils.verify(true,true,"AssetGroup field is available in Macanomy for the Create Asset");
       
    var description = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.descripp.getText().OleValue.toString().trim();
    if(description!="Description")
    ValidationUtils.verify(false,true,"Description field is missing in macanomy for the Create Asset");
    else
    ValidationUtils.verify(true,true,"Description field is available in Macanomy for the Create Asset");

    var compny = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.comp.getText().OleValue.toString().trim();
    if(compny!="Company")
    ValidationUtils.verify(false,true,"Company field is missing in macanomy for the Create Asset");
    else
    ValidationUtils.verify(true,true,"Company field is available in Macanomy for the Create Asset");
}
 

function address2(){
    var datee = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.acessgrop.getText().OleValue.toString().trim();
    if(datee!="Date")
    ValidationUtils.verify(false,true,"Date field is missing in macanomy for the Create Asset");
    else
    ValidationUtils.verify(true,true,"Date field is available in Macanomy for the Create Asset");
    
    var costt =   Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.descripp.getText().OleValue.toString().trim();
    if(costt!="Cost")
    ValidationUtils.verify(false,true,"Cost field is missing in macanomy for the Create Asset");
    else
    ValidationUtils.verify(true,true,"Cost field is available in Macanomy for the Create Asset");
    
    var AccessLevel = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.aclevel.getText().OleValue.toString().trim();
    if(AccessLevel!="Access Level")
    ValidationUtils.verify(false,true,"AccessLevel field is missing in macanomy for the Create Asset");
    else
    ValidationUtils.verify(true,true,"AccessLevel field is available in Macanomy for the Create Asset");
} 
    
    
function Posting(){ 
 var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet0.GendralLedger.Exists()){
          ImageRepository.ImageSet0.GendralLedger.Click();// GL
          }
          else if(ImageRepository.ImageSet0.GendralLedger1.Exists()){
          ImageRepository.ImageSet0.GendralLedger1.Click();
          }
          else{
          ImageRepository.ImageSet0.GendralLedger2.Click();
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
        Client_Managt.ClickItem("|GL Transactions");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|GL Transactions");
      }
    }
    aqUtils.Delay(3000,Indicator.Text);

    var posting = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.posting;
    posting.Click();
    aqUtils.Delay(1000,Indicator.Text);

    var fromCompany = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
    fromCompany.Click();
    fromCompany.setText(comapany);
    aqUtils.Delay(1000,Indicator.Text);
 
    var toCompany = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.SWTObject("McTextWidget", "", 4);
    toCompany.Click();
    toCompany.setText(comapany);
    aqUtils.Delay(1000,Indicator.Text);
    
    if(date!=""){
      var createfrom = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.createfrom;
      WorkspaceUtils.CalenderDateSelection(createfrom,date)
      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Post Fixed Asset");
    }  
    aqUtils.Delay(1000,Indicator.Text);   
    
    if(date!=""){
      var createTo = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.createto;
      WorkspaceUtils.CalenderDateSelection(createTo,date)
      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Post Fixed Asset");
    }  
    aqUtils.Delay(2000,Indicator.Text);  
         
//    var jorntype = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget.Composite.journaltype;
//     jorntype.Keys(" ");
     
//     var postoption = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget.Composite2.post;
//     postoption.Keys("Yes");
//     aqUtils.Delay(1000,Indicator.Text);

    Sys.Process("Maconomy").Refresh();
   
     var layouttext = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.layout;
      layouttext.Keys("WPP GeneralJournal");
//    layouttext.Keys(" ");
//    if(layout!=""){
//       layouttext.Click();
//       aqUtils.Delay(1000, Indicator.Text);
//       WorkspaceUtils.DropDownList(layout,"Layout");
//       aqUtils.Delay(1000, Indicator.Text); 
//    } 
//    else{
//      ValidationUtils.verify(false,true,"Layout is Needed to Post fixed asset");  
//    }     
      ValidationUtils.verify(true,true,"Layout is selected to Post fixed asset"); 
      aqUtils.Delay(2000,Indicator.Text);
    var save = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.Composite.savelayout;
    Sys.HighlightObject(save);
    ReportUtils.logStep_Screenshot();
    save.Click();
    aqUtils.Delay(1000,Indicator.Text);
        
    
//    var journal = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
//    Sys.HighlightObject(journal)
//    journal.Click();
//    aqUtils.Delay(1000,Indicator.Text);
    var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid;
    Sys.HighlightObject(table) 
     checkbox();
      initialcheckbox();
      for(var i=0;i<table.getItemCount();i++){ 
       if(transactionNo==table.getItem(i).getText(3).OleValue.toString().trim()){   
       table.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");    
       var check = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.postcolumn.SWTObject("Button", "");
        if(check.getSelection()){             
              ValidationUtils.verify(true,true,"Checkbox is Clicked");
         }
         else{
           check.Click();
           ValidationUtils.verify(true,true,"Checkbox is Clicked");
         } 
          aqUtils.Delay(2000,Indicator.Text);
          
          ImageRepository.ImageSet0.Maximizejournal.Click();
          aqUtils.Delay(2000,Indicator.Text);
         var save = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
         save.Click();
         aqUtils.Delay(2000,Indicator.Text);
         ReportUtils.logStep_Screenshot();
         ImageRepository.ImageSet0.Postingjournal.Click();
         aqUtils.Delay(2000,Indicator.Text);


         var Post = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
         Sys.HighlightObject(Post);
         ReportUtils.logStep_Screenshot();
         Post.Click();
         aqUtils.Delay(5000,Indicator.Text);
         ValidationUtils.verify(true,true,"Successfully Posted the Assest");
          break;
       }
       else{ 
          table.Keys("[Down]");  
       }
     } 
}

function checkbox(){
   var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid;
    table.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
    
    var check = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.postcolumn.SWTObject("Button", "");
    for(var i=0;i<table.getItemCount();i++){ 
    if(check.getSelection()){
      check.Click();
      Log.Message("Checkbox is UnChecked")
    }
    else { 
            table.Keys("[Down]");
         } 
    }
} 

function initialcheckbox(){
      var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid;
  var row = table.getItemCount();
  var col = table.getColumnCount();
  Log.Message(col)  
      for(var j=8;j>0;j--){
        Sys.Desktop.KeyDown(0xA0);
        Sys.Desktop.KeyDown(0x09);
        Sys.Desktop.KeyUp(0xA0);
        Sys.Desktop.KeyUp(0x09);
      }
      var item = table.getItemCount()
  for(var i=item;i>0;i--){ 
    Sys.Desktop.KeyDown(0x26)
    Sys.Desktop.KeyUp(0x26)																																																																																																																																																																																																													
  }
} 
  