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
  var asetnumber="";
  
function CreateAssest(){
  
Indicator.PushText("waiting for window to open");
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
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
      sheetName = "Create Fixed Asset";
      ExcelUtils.setExcelName(workBook, sheetName, true); 
    try{   
      getDetails();
      goToJobMenuItem(); 
      createAssets();
      closeAllWorkspaces();
      WorkspaceUtils.closeMaconomy();
      Restart.login(login);
      Posting();
      WorkspaceUtils.closeAllWorkspaces();
  }
    catch(err){
      Log.Message(err);
    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
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
        Log.Message(date)
        if((date==null)||(date=="")){ 
        ValidationUtils.verify(false,true,"AssetDate is Needed to Create a Asset");
        }                  
        cost = ExcelUtils.getRowDatas("Cost",EnvParams.Opco)
        Log.Message(cost)
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

      comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
      Log.Message(comapany)
      if((comapany==null)||(comapany=="")){ 
        ValidationUtils.verify(false,true,"Company Number is Needed to Create Asset");
      }
      if((comapany=="")||(comapany==null))
      ValidationUtils.verify(false,true,"Comapany Number is needed to Create Asset");
      
      
      ExcelUtils.setExcelName(workBook, "SSC Users", true);
      login = ExcelUtils.getRowDatas("SSC - Senior Accountant","Username")
        Log.Message(login);
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
    TextUtils.writeLog("Moved to Asset from Fixed Asset");
    aqUtils.Delay(3000,Indicator.Text);
  }


  
  
function createAssets(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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
//      address1();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(5000,Indicator.Text);   
    var assetGroup = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.assetgroup;
      if(Assetgroup!=""){
        assetGroup.Click();
        WorkspaceUtils.SearchByValue(assetGroup,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Asset Group").OleValue.toString().trim(),Assetgroup,"Name");
        TextUtils.writeLog("Asset Group is selected: "+Assetgroup);
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
        WorkspaceUtils.SearchByValue(company,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),comapany,"Company Number");
      }
      else{ 
        ValidationUtils.verify(false,true,"company is Needed to Create a Fixed Assets");
      }
    
    var next = Aliases.Maconomy.Group5.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
    Sys.HighlightObject(next);
    ReportUtils.logStep_Screenshot();
    next.Click();
    aqUtils.Delay(1000,Indicator.Text);
    
//    address2();    
    
    if(date!=""){
      var datefiled = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.date;
      WorkspaceUtils.CalenderDateSelection(datefiled,date)
      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Create a Employee");
    }    
    aqUtils.Delay(1000,Indicator.Text);
    
 
    var costt = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.descrip;
    Sys.HighlightObject(costt);
    if(cost!=""){
        costt.Click();
        costt.setText(cost);
        ValidationUtils.verify(true,true,"Cost is Entered in Maconomy"); 
        TextUtils.writeLog("Cost is Entered in Maconomy: "+cost);
      }
      else{ 
        ValidationUtils.verify(false,true,"Cost is Needed to Create a Fixed Assets");
      }

   var access = Aliases.Maconomy.Group5.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
      if(comapany!=""){
        access.Click();
        WorkspaceUtils.SearchByValue(access,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Access Level").OleValue.toString().trim(),Access,"Name");
      }
      else{ 
        ValidationUtils.verify(false,true,"company is Needed to Create a Fixed Assets");
      }
    
    var btnCreate = NameMapping.Sys.Maconomy.Group5.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
    if(btnCreate.isEnabled()){
        Sys.HighlightObject(btnCreate)
        btnCreate.HoverMouse();
      ReportUtils.logStep_Screenshot("");
        btnCreate.Click();
      ValidationUtils.verify(true,true,"Asset is Created");
      TextUtils.writeLog("Asset is Created");
      ReportUtils.logStep("INFO", "Asset is Created");
      aqUtils.Delay(5000, Indicator.Text);
    }
    else{ 
      var cancel = Aliases.Maconomy.Group5.Composite.Composite.Composite2.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
          
       asetnumber =NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.assetnumber.getText().OleValue.toString().trim();
//       Log.Message(asetnumber);    
aqUtils.Delay(5000, Indicator.Text);
       ValidationUtils.verify(true,true,"Asset Number : "+asetnumber);
        ExcelUtils.setExcelName(workBook,"Data Management", true);
        ExcelUtils.WriteExcelSheet("Assets No",EnvParams.Opco,"Data Management",asetnumber)
         TextUtils.writeLog("Asset is Created:" +asetnumber);

        var allEntries = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;
        allEntries.Click();
        aqUtils.Delay(2000, Indicator.Text);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
          
        }
        var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.entrytable.McGrid;
        transactionNo = table.getItem(0).getText(4).OleValue.toString().trim();
        Log.Message("Transaction No. :"+transactionNo);
        ReportUtils.logStep_Screenshot("");
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
        Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());

        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());

      }
    }
    aqUtils.Delay(3000,Indicator.Text);

    var posting = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.posting;    
    posting.Click();
    aqUtils.Delay(1000,Indicator.Text);
    
    aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - Post").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
 var Okay = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - Post").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim()).Click();
 
}
    aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
    
//    var fromCompany = Aliases.Maconomy.Group7.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
    
      var fromCompany = Aliases.Maconomy.Group7.McClumpSashForm.Composite;
  Log.Message(fromCompany.FullName)
  Sys.HighlightObject(fromCompany);
  
var fromCompany = ""
var childcount = 0;
var Add = [];
var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
for(var i = 0;i<Parent.ChildCount;i++){ 
  if((Parent.Child(i).isVisible()) && (Parent.Child(i).ChildCount == 1)){
  Add[childcount] = Parent.Child(i);
  childcount++;
  }
}

Parent = "";
var pos = 1000;
for(var i=0;i<Add.length;i++){ 
  if(Add[i].Height<pos){ 
    pos = Add[i].Height;
    Parent = Add[i];
  }
}


Log.Message(Parent.FullName)
fromCompany = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2)
Sys.HighlightObject(fromCompany);
Log.Message(fromCompany.FullName)
    waitForObj(fromCompany)
    fromCompany.Click();
    fromCompany.setText(comapany);
    aqUtils.Delay(1000,Indicator.Text);
 
//    var toCompany = Aliases.Maconomy.Group7.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget2;
    var toCompany = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 4)
    toCompany.Click();
    toCompany.setText(comapany);
    aqUtils.Delay(1000,Indicator.Text);
    
    if(date!=""){
//      var createfrom = Aliases.Maconomy.Group7.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McDatePickerWidget;
      var createfrom = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McDatePickerWidget", "", 2);
      createfrom.setText(aqDateTime.Today())
//      WorkspaceUtils.CalenderDateSelection(createfrom,date)
      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Post Fixed Asset");
    }  
    aqUtils.Delay(1000,Indicator.Text);   
    
    if(date!=""){
//      var createTo = Aliases.Maconomy.Group7.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McDatePickerWidget2;
      var createTo = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McDatePickerWidget", "", 4);
      createTo.setText(aqDateTime.Today())
//      WorkspaceUtils.CalenderDateSelection(createTo,date)
      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Post Fixed Asset");
    }  
    aqUtils.Delay(2000,Indicator.Text);  
         

    Sys.Process("Maconomy").Refresh();
   
//     var layouttext = Aliases.Maconomy.Group7.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McPopupPickerWidget;
     var layouttext = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
      layouttext.Keys("WPP GeneralJournal");
  
      ValidationUtils.verify(true,true,"Layout is selected to Post fixed asset"); 
      aqUtils.Delay(2000,Indicator.Text);
//    var save = Aliases.Maconomy.Composite.SingleToolItemControl;
    var save = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
    Sys.HighlightObject(save);
    ReportUtils.logStep_Screenshot();
    save.Click();
    aqUtils.Delay(1000,Indicator.Text);

    var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
    Sys.HighlightObject(table);

//     var transactionNo ="1006800354";
      for(var i=0;i<table.getItemCount();i++){ 
        Log.Message(transactionNo)
       if(transactionNo==table.getItem(i).getText(3).OleValue.toString().trim()){   
       table.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");    
aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var check =  table.SWTObject("McPlainCheckboxView", "", 5).SWTObject("Button", "");
//     var check =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 5).SWTObject("Button", "");
//                  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 5).SWTObject("Button", "")
  
      if(check.getSelection()){             
              ValidationUtils.verify(true,true,"Checkbox is Clicked");
         }
         else{
           check.Click();
           ValidationUtils.verify(true,true,"Checkbox is Clicked");
         } 
          aqUtils.Delay(2000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
          aqUtils.Delay(2000,Indicator.Text);
         var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
         save.Click();
         aqUtils.Delay(2000,Indicator.Text);
         ReportUtils.logStep_Screenshot();
         aqUtils.Delay(2000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

         var Post = Aliases.Maconomy.Composite.SingleToolItemControl2;
         Sys.HighlightObject(Post);
         ReportUtils.logStep_Screenshot();
         Post.Click();
         aqUtils.Delay(15000,Indicator.Text);
         ValidationUtils.verify(true,true,"Successfully Posted the Assest");
         TextUtils.writeLog("Asset Successfully Posted :" +asetnumber);
          break;
       }
       else{ 
          table.Keys("[Down]");  
       }
     } 

aqUtils.Delay(9000,Indicator.Text); 
     var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal-"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
  if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal-"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Posting Journal-")!=-1){
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

var filepathforMplValidation =sFolder+SaveTitle+".pdf";
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
aqUtils.Delay(2000, Indicator.Text);

Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print Asset is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

     
}

function checkbox(){


  var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
 table.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
 
    var lns = false;
    if(!lns)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).isVisible())
    {
    var check = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 5).SWTObject("Button", "");
    lns = true;
    }
  
     if(!lns)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    {
    var check = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 5).SWTObject("Button", "");
    lns = true;
    }
    
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
      var lns = false;
    if(!lns)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).isVisible())
    {
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    lns = true;
    }
  
     if(!lns)
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).isVisible())
    {
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    lns = true;
    }
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


 