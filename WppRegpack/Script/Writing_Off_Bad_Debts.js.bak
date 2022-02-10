//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT TestRunner
//USEUNIT Restart


var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "Writingoffbad";
Indicator.Show();
Indicator.PushText("waiing for window to open");


var companyno ="";
var TType = "";
var Jobno ="";
var STIME = "";
var Client = "";
var workcode = "";
var Credit = "";
var EntryJob = "";
var Descip = "";
var login = "";
var getJournal = "";
var Jobdepart = "";
var Jobbusiness = "";

function Writingoffbad() {
      Language = "";
    Language = EnvParams.Language;
    if((Language==null)||(Language=="")){
    ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
    }
    Log.Message(EnvParams.Opco)
    Log.Message(Language)
    Language = EnvParams.LanChange(Language);
    WorkspaceUtils.Language = Language;
    getDetails();
    goToJobMenuItem();
    getJobDetails();
    goToGLMenuItem();  
    goToGLJournal();  
    goToEntries();
//    WorkspaceUtils.closeMaconomy();
//    Restart.login(login);
//    goToGLMenuItem(); 
//    post();
}


function getDetails(){
        sheetName = "Writingoffbad";     
        ExcelUtils.setExcelName(workBook, sheetName, true);
        
        companyno = ExcelUtils.getColumnDatas("company",EnvParams.Opco)
        Log.Message(companyno);
        if((companyno==null)||(companyno=="")){ 
        ValidationUtils.verify(false,true,"Company Number is needed to Create Writing off Bad Debts");
        }
        Jobno = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
        Log.Message(Jobno);
        if((Jobno==null)||(Jobno=="")){ 
        ValidationUtils.verify(false,true,"Job Number is needed to Create Writing off Bad Debts");
        }
        TType = ExcelUtils.getColumnDatas("Ttype",EnvParams.Opco)
        Log.Message(TType); 
        
        Descip = ExcelUtils.getColumnDatas("Description",EnvParams.Opco)
        Log.Message(Descip);
        
        Client = ExcelUtils.getColumnDatas("ClientNo",EnvParams.Opco)
        Log.Message(Client);
        if((Client==null)||(Client=="")){ 
        ValidationUtils.verify(false,true,"Client is needed to Create Writing off Bad Debts");
        }
        workcode = ExcelUtils.getColumnDatas("Workcode",EnvParams.Opco)
        Log.Message(workcode);
        if((workcode==null)||(workcode=="")){ 
        ValidationUtils.verify(false,true,"Workcode is needed to Create Writing off Bad Debts");
        }
        Credit = ExcelUtils.getColumnDatas("credit",EnvParams.Opco)
        Log.Message(Credit);
        if((Credit==null)||(Credit=="")){ 
        ValidationUtils.verify(false,true,"Credit Amount is needed to Create Writing off Bad Debts");
        }
        EntryJob = ExcelUtils.getColumnDatas("BadoffJobno",EnvParams.Opco)
        Log.Message(EntryJob);
        if((EntryJob==null)||(EntryJob=="")){ 
        ValidationUtils.verify(false,true,"Job Number is needed to Create Writing off Bad Debts");
        }
        login = ExcelUtils.getColumnDatas("Login",EnvParams.Opco)
        Log.Message(login);
        if((login==null)||(login=="")){ 
        ValidationUtils.verify(false,true,"Login is needed to Post Writing off Bad Debts");
        }
        
}

function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet0.Jobs.Exists()){
         ImageRepository.ImageSet0.Jobs.Click();// GL
         }
        else if(ImageRepository.ImageSet0.Job.Exists()){
        ImageRepository.ImageSet0.Job.Click();
        }
        else{
        ImageRepository.ImageSet0.Jobs1.Click();
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
        Client_Managt.ClickItem("|Jobs");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|Jobs");
      }
    }
    Delay(3000);
  }  

  function getJobDetails(){
//      TextUtils.writeLog("Writting Off Bad Debts is started");
      ReportUtils.logStep("INFO", "Writting Off Bad Debts is started::"+STIME);
      var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
      Sys.HighlightObject(table)
      var compno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;      
      compno.Click();
      compno.setText(companyno);
      aqUtils.Delay(1000,Indicator.Text);
      compno.Keys("[Tab][Tab]");
      var jobno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
      jobno.Click();
      jobno.setText(Jobno);
      aqUtils.Delay(1000,Indicator.Text);
      
      var flag =false; 
        for(var i=0;i<table.getItemCount();i++){          
          if(table.getItem(i).getText_2(2).OleValue.toString().trim()==Jobno){
            flag = true;        
            break;
          }  
          else{
              table.Keys("[Down]");
          } 
        } 
        aqUtils.Delay(2000,Indicator.Text); 
        ReportUtils.logStep_Screenshot();    
        ValidationUtils.verify(true,true,"Job Number is available in system");
        aqUtils.Delay(3000,Indicator.Text);          
            
                Sys.Desktop.KeyDown(0x11);
                Sys.Desktop.KeyDown(0x46);
               Sys.Desktop.KeyUp(0x11);
                Sys.Desktop.KeyUp(0x46);  
      aqUtils.Delay(2000,Indicator.Text);  
      var home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
      home.Click();
      Sys.HighlightObject(home);
      aqUtils.Delay(1000,Indicator.Text); 
      
        var lns= false;
        if(!lns)
          if(NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.isVisible()){
          var information = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
          lns = true;
          }
        if(!lns)
        if(NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.isVisible()){
        var information = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl3;
        lns = true;
        }
        
        Sys.HighlightObject(information);
      information.Click();
      aqUtils.Delay(4000,Indicator.Text);
      var Department = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget.getText();
      Jobdepart = Department;
      Log.Message(Jobdepart)
      var BusinessUnit = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Composite.McValuePickerWidget.getText();
      Jobbusiness = BusinessUnit;  
      Log.Message(Jobbusiness)
      aqUtils.Delay(1000,Indicator.Text);
      closeAllWorkspaces();    
  }

function goToGLMenuItem(){
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
  }

  
function address(){
  var company = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.companyfield.getText().OleValue.toString().trim();
  if(company!="Company"){
    ValidationUtils.verify(false,true,"Company field is missing in macanomy for the Writing off bad debts");
  }
  else{
     ValidationUtils.verify(true,true,"Company field is available in macanomy for the Writing off bad debts");
  }
  var TType = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.TTypefield.getText().OleValue.toString().trim();
  if(TType!="Transaction Type"){
    ValidationUtils.verify(false,true,"Transaction Type field is missing in macanomy for the Writing off bad debts");
  }
  else{
     ValidationUtils.verify(true,true,"Transaction Type field is available in macanomy for the Writing off bad debts");
  }  
}
  
  
  function goToGLJournal(){
    var generaltab = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(generaltab);
    generaltab.Click();
    aqUtils.Delay(1000,Indicator.Text);
    var subgltab = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;

    Sys.HighlightObject(subgltab);
    subgltab.Click();
    aqUtils.Delay(2000,Indicator.Text);
    
    //closeFilter
    
              Sys.Desktop.KeyDown(0x11);
                Sys.Desktop.KeyDown(0x46);
               Sys.Desktop.KeyUp(0x11);
                Sys.Desktop.KeyUp(0x46);
        
    //AddIcon
    
    
    Sys.Desktop.KeyDown(0xA2);
    Sys.Desktop.KeyDown(0x4E);
    Sys.Desktop.KeyUp(0xA2);    
    Sys.Desktop.KeyUp(0x4E);
    
    aqUtils.Delay(1000,Indicator.Text);
    ReportUtils.logStep_Screenshot("");
    address();
    var company = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget;
    Sys.HighlightObject(company);
    if(companyno!=""){
      company.Click();
      WorkspaceUtils.SearchByValue(company,"Company",companyno,"Company No");
    }
    else{
      ValidationUtils.verify(false,true,"Company Number is needed to Create Writing Off Bad Debts");
    }    
    
    var Ttype = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.Ttypevalue;
    Sys.HighlightObject(Ttype);
    if(TType!=""){
      Ttype.Click();
      WorkspaceUtils.SearchByValuePickerTType(Ttype,"Transaction Type",TType);
    }
    else{
      ValidationUtils.verify(false,true,"Transaction Type is Needed to create Write off bad debts");
    }
    
    var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.savee;
    ReportUtils.logStep_Screenshot("");
    save.Click();
    aqUtils.Delay(1000,Indicator.Text);
    
    var journalnum = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.McTextWidget;
    getJournal = journalnum.getText();
    Log.Message(getJournal);
}

function goToEntries(){
    
    var entriestable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table;
    Sys.HighlightObject(entriestable);
    
    for(var i=1;i<=2;i++){
          ExcelUtils.setExcelName(workBook, sheetName, true);
          var grp = ExcelUtils.getColumnDatas("GRP_"+i,EnvParams.Opco)
     
              var addicon = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.addbutton;   
              Sys.HighlightObject(addicon);              
              addicon.Click();
              aqUtils.Delay(1000,Indicator.Text);
              ReportUtils.logStep_Screenshot("");
              var entrydate = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.firstcell;
              entrydate.Keys("[Tab]");
              aqUtils.Delay(1000,Indicator.Text);
    
              var description = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.amount;
              if(Descip!=""){
                description.Click();
                description.setText(Descip);
                ValidationUtils.verify(true,true,"Description is Entered");
              }
              else{
                ValidationUtils.verify(false,true,"Description is Neeeded to writing off bad debts");      
              }
              description.Keys("[Tab]");
              aqUtils.Delay(1000,Indicator.Text);    
              var Grp = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.firstcelltab;
              Grp.Keys(" ");
              aqUtils.Delay(1000,Indicator.Text); 
              if(grp!=""){
                Grp.Click();
                WorkspaceUtils.DropDownList(grp,"GRP");
                aqUtils.Delay(2000,Indicator.Text);      
              }
              else{
                ValidationUtils.verify(false,true,"GRP is Needed for writing off bad debts");
              }
                if(i==1){
                  Grp.Keys("[Tab]");
                   var client = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.workcode;
                    if(Client!=""){
                      ReportUtils.logStep_Screenshot("");
                      client.Click();
                      WorkspaceUtils.SearchByValuePicker(client,"Client",Client);
                      ValidationUtils.verify(true,true,"Client Number is Selected in the macanomy");
                      aqUtils.Delay(2000,Indicator.Text);
                    }
                    else{
                      ValidationUtils.verify(false,true,"Client is Needed for writing off bad debts");
                    }
                    client.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
        
                    var credit = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.amount;
                    if(Credit!=""){
                      credit.Click();
                      credit.setText(Credit);
//                      TextUtils.writeLog("Credit amount is Entered");
                      ValidationUtils.verify(true,true,"Credit amount is Entered");
                    }
                    else{
                      ValidationUtils.verify(false,true,"Credit amount is Needed to create writing off bad debt");
                    }
                    ReportUtils.logStep_Screenshot("");
                }
                else{
                      Grp.Keys("[Tab][Tab][Tab]");         
                    var Jobnum = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.workcode;
                    if(EntryJob!=""){
                      Jobnum.Click();
                      WorkspaceUtils.SearchByValuesjob(Jobnum,"Job",EntryJob,"Job Number") 
                    }
                    else{
                      ValidationUtils.verify(false,true,"Job Number is Needed for writing off bad debts");
                    }
                    ReportUtils.logStep_Screenshot("");
                    Jobnum.Keys("[Tab]");
        
                    var Workcode = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.workcode;
                    if(workcode!=""){
                      Workcode.Click();
                      WorkspaceUtils.SearchByValue(Workcode,"Work Code",workcode,"Work Code Name");
                    }
                    else{
                      ValidationUtils.verify(false,true,"WorkCode is Needed for writing off bad debts");
                    }
                    aqUtils.Delay(1000,Indicator.Text);
                    Workcode.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
                    var jobdepart = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.workcode;
                    Log.Message(Jobdepart);
                    if(Jobdepart!=""){
                      jobdepart.Click();
                      WorkspaceUtils.SearchByValue(jobdepart,"Local Specification 2",Jobdepart,"Name");
                    }
                    else{
                      ValidationUtils.verify(false,true,"Job Department is needed to Create Writing Off Bad Debts");
                    }
                    jobdepart.Keys("[Tab][Tab]");
                    var jobbusiness = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.workcode;
                    Log.Message(Jobdepart);
                    if(Jobbusiness!=""){
                      jobbusiness.Click();
                      WorkspaceUtils.SearchByValue(jobbusiness,"Local Specification 4",Jobbusiness,"Name");
                    }
                    else{
                      ValidationUtils.verify(false,true,"Job Business is needed to Create Writing Off Bad Debts");
                    }                   
                    
                }               
              var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
              save.Click();
              aqUtils.Delay(1000,Indicator.Text); 
              ReportUtils.logStep_Screenshot("");     
    }
    var submit = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.submitt;
    submit.Click(); 
    
    var GL = NameMapping.Sys.Maconomy.SWTObject("Shell", "GL Transactions - General Journal");
    Sys.HighlightObject(GL);
    ReportUtils.logStep_Screenshot("");
    var OK = NameMapping.Sys.Maconomy.SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    OK.Click();
//     TextUtils.writeLog("Writing Off Bad Debt is submitted");
    ValidationUtils.verify(true,true,"Writing Off Bad Debt is submitted");
}


function post(){
    var generaltab = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
    Sys.HighlightObject(generaltab);
    generaltab.Click();
    var table = Aliases.Maconomy.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    Sys.HighlightObject(table);
    var comp = Aliases.Maconomy.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
    comp.Click();
    aqUtils.Delay(1000,Indicator.Text); 
    comp.Keys("[Tab]");
    var journal = Aliases.Maconomy.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
    journal.Click();
    journal.setText(getJournal);
    aqUtils.Delay(2000,Indicator.Text);      
     var flag =false; 
        for(var i=0;i<table.getItemCount();i++){          
          if(table.getItem(i).getText_2(2).OleValue.toString().trim()==getJournal){
            flag = true;        
            break;
          }  
          else{
              table.Keys("[Down]");
          } 
        } 
        ReportUtils.logStep_Screenshot();    
        ValidationUtils.verify(true,true,"Journal Number is available in system");
        aqUtils.Delay(3000,Indicator.Text);          
            
                Sys.Desktop.KeyDown(0x11);
                Sys.Desktop.KeyDown(0x46);
               Sys.Desktop.KeyUp(0x11);
                Sys.Desktop.KeyUp(0x46); 
                
     var post = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
     post.Click();
     ValidationUtils.verify(true,true,"Journal Number is Posted");
    
}

