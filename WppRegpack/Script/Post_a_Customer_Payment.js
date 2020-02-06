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
var sheetName = "Post a Customer Payment";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");
  
var JournalNo ="";
  
function CustomerPayment() { 
  
Indicator.PushText("waiting for window to open");
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  aqUtils.Delay(3000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Combined Biller, IC & AR","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);  
}


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
      getDetails();
    goToARMenuItem();
    gotoPost();
    closeAllWorkspaces();
    goToJobMenuItem();
    invoicejob();
    closeAllWorkspaces();
    goToGeneralLedger();
    GLLookups();
    WorkspaceUtils.closeAllWorkspaces();    
}



function getDetails(){
        ExcelUtils.setExcelName(workBook, sheetName, true);       
        sheetName = "Post a Customer Payment";   
        
        ExcelUtils.setExcelName(workBook, sheetName, true);
        JournalNo = ExcelUtils.getRowDatas("Journal No",EnvParams.Opco)
        if((JournalNo=="")||(JournalNo==null)){
              ExcelUtils.setExcelName(workBook, "Data Management", true);
              JournalNo = ReadExcelSheet("Journal No",EnvParams.Opco,"Data Management");
        }   
               
} 

function goToARMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet0.Account_Receivable.Exists()){
         ImageRepository.ImageSet0.Account_Receivable.Click();// GL
         }
        else if(ImageRepository.ImageSet0.Acc_Receivable_1.Exists()){
        ImageRepository.ImageSet0.Acc_Receivable_1.Click();
        }
        else{
        ImageRepository.ImageSet0.Acc_Receivable_2.Click();
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
        Client_Managt.ClickItem("|AR Transactions");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|AR Transactions");
      }
    }
    Delay(3000);
  }


function gotoPost(){ 
//  TextUtils.writeLog("Post Customer Payment is started");
        ReportUtils.logStep("INFO", "Post Customer Payment is started::"+STIME);
        var client = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.clientpayment;
        Sys.HighlightObject(client);
        var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.table.clienttable;
        Sys.HighlightObject(table);
        var firstcell = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.table.clienttable.firstcell;
        firstcell.Click();
        firstcell.setText(JournalNo);
        aqUtils.Delay(1000,Indicator.Text);
        var flag =false;
        for(var i=0;i<table.getItemCount();i++){
          if((table.getItem(i).getText_2(0).OleValue.toString().trim()==JournalNo)){
          flag = true;
            break;
          } 
          else{
            table.Keys("[Down]");
          } 
        } 
        ValidationUtils.verify(true,true,"Journal Number is available in system");
        aqUtils.Delay(1000,Indicator.Text);
        ReportUtils.logStep_Screenshot("");
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
               
        
        ReportUtils.logStep("INFO", "Journal Number is listed in the table");            
        aqUtils.Delay(3000,Indicator.Text);       
        ReportUtils.logStep_Screenshot("");
        var post = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.post;
        post.Click();
        ReportUtils.logStep_Screenshot("");          
        ValidationUtils.verify(true,true,"Journal is Posted");        
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

function invoicejob(){  
  
        ExcelUtils.setExcelName(workBook, sheetName, true);       
        sheetName = "Post a Customer Payment";  
  
        companyno = ExcelUtils.getRowDatas("company",EnvParams.Opco)
        if((companyno==null)||(companyno=="")){ 
        ValidationUtils.verify(false,true,"Company Number is needed to Check Invoice status"); 
        }
        Jobno = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
          if((Jobno=="")||(Jobno==null)){
            ExcelUtils.setExcelName(workBook, "Data Management", true);
            Jobno = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
          }  
        if((Jobno=="")||(Jobno==null))
        ValidationUtils.verify(false,true,"Job Number is needed to Check Invoice status"); 
        
      var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid      
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
                        
       var invoice = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.invoice;
       Sys.HighlightObject(invoice);       
       ReportUtils.logStep_Screenshot(""); 
       invoice.Click();
       aqUtils.Delay(1000,Indicator.Text);
       
       var invoicehistory = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.history;
       Sys.HighlightObject(invoicehistory);
       invoicehistory.Click();
       aqUtils.Delay(1000,Indicator.Text);
       
       var invoicetable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.invoicetable_1;
       Sys.HighlightObject(invoicetable);
       ReportUtils.logStep_Screenshot(""); 
       var row = invoicetable.getItemCount();
       var column = invoicetable.getColumnCount();
       
         for(var i=0;i<invoicetable.getItemCount();i++){         
                 if(invoicetable.getItem(i).getText(8).OleValue.toString().trim()=="Reconciled"){  
                  break;
                 }
                 else{
                     invoicetable.Keys("[Down]");
                 }       
      }
      ReportUtils.logStep_Screenshot(""); 
      //  TextUtils.writeLog("Payment status changed as Reconciled");
      ValidationUtils.verify(true,true,"Payment status changed as Reconciled");
      
}


function goToGeneralLedger(){
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
        Client_Managt.ClickItem("|GL Lookups");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|GL Lookups");
      }
    }
    Delay(3000);
  }

  
function GLLookups(){
  var journal = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.journal;
  journal.Click();
  aqUtils.Delay(2000,Indicator.Text);
  ReportUtils.logStep_Screenshot("");
  var postedjournal = NameMapping.Sys.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(postedjournal) ;
    aqUtils.Delay(1000,Indicator.Text);
  var all = NameMapping.Sys.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.AllButton;
   all.Click();
   var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
   
   var journalnum = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
   journalnum.Click();
   journalnum.setText(JournalNo);
        aqUtils.Delay(1000,Indicator.Text);
        var flag =false;
        for(var i=0;i<table.getItemCount();i++){
          if((table.getItem(i).getText_2(0).OleValue.toString().trim()==JournalNo)){
          flag = true;
            break;
          } 
          else{
            table.Keys("[Down]");
          } 
        } 
        aqUtils.Delay(1000,Indicator.Text);
        ReportUtils.logStep_Screenshot("");
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
        aqUtils.Delay(3000,Indicator.Text);        
        ValidationUtils.verify(flag,true,"Posted Journal Number is available in the List");
        ReportUtils.logStep("INFO", "Posted Journal Number is available in the List");            
        aqUtils.Delay(1000,Indicator.Text);       
        ReportUtils.logStep_Screenshot("");
      
}
  

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}



