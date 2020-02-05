//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT TestRunner

 
var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "AR Single Payment";
Indicator.Show();
Indicator.PushText("waiing for window to open");


var Jobno ="";
var companyno ="";
var IvoiceNumber ="";
var Descip="";
var invoiceamount ="";
var currency ="";
var clientnum = "";
var TP ="";
var STIME = "";
var Clientbalance ="";

function AccountsReceivable() {
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
    invoicejob();
    closeAllWorkspaces();
    goToARMenuItem();    
    gotoReceivable();
}

function getDetails(){

         sheetName = "AR Single Payment";     
        ExcelUtils.setExcelName(workBook, sheetName, true);
        
        companyno = ExcelUtils.getRowDatas("company",EnvParams.Opco)
        if((companyno==null)||(companyno=="")){ 
        ValidationUtils.verify(false,true,"Company Number is needed to Create Single Invoice"); 
        }  
//        IvoiceNumber = ExcelUtils.getRowDatas("Ivoice Number",EnvParams.Opco)
//        if((IvoiceNumber==null)||(IvoiceNumber=="")){ 
//        ValidationUtils.verify(false,true,"Ivoice Number is needed to Create Single Invoice"); 
//        }    
        Descip = ExcelUtils.getRowDatas("Description",EnvParams.Opco)
        if((Descip==null)||(Descip=="")){ 
        ValidationUtils.verify(false,true,"Description is needed to Create Single Invoice"); 
        }  
        currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
        if((currency==null)||(currency=="")){ 
        ValidationUtils.verify(false,true,"Currency is needed to Create Single Invoice"); 
        }   
        TP = ExcelUtils.getRowDatas("TP",EnvParams.Opco)
        if((TP==null)||(TP=="")){ 
        ValidationUtils.verify(false,true,"TP is needed to Create Single Invoice"); 
        }  
        Jobno = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
          if((Jobno=="")||(Jobno==null)){
            ExcelUtils.setExcelName(workBook, "Data Management", true);
            Jobno = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
          }  
        if((Jobno=="")||(Jobno==null))
        ValidationUtils.verify(false,true,"Job Number is needed to Create Single Invoice");
        
        clientnum = ExcelUtils.getRowDatas("Clientno",EnvParams.Opco)
          if((clientnum=="")||(clientnum==null)){
            ExcelUtils.setExcelName(workBook, "Data Management", true);
            clientnum = ReadExcelSheet("Clientno",EnvParams.Opco,"Data Management");
          }  
        if((clientnum=="")||(clientnum==null))
        ValidationUtils.verify(false,true,"Client Number is needed to Create Single Invoice");


            
          
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
      ReportUtils.logStep("INFO", "Customer Payment for Single Invoice is started::"+STIME);
      var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid      
      var compno = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
      compno.Click();
      compno.setText(companyno);
      aqUtils.Delay(2000,Indicator.Text);
      compno.Keys("[Tab][Tab]");
      var jobno = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
      jobno.Click();
      jobno.setText(Jobno);
      aqUtils.Delay(2000,Indicator.Text);
      
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
        aqUtils.Delay(3000,Indicator.Text); 
        ReportUtils.logStep_Screenshot();    
        ValidationUtils.verify(true,true,"Job Number is available in system");
        aqUtils.Delay(3000,Indicator.Text);          
            
                Sys.Desktop.KeyDown(0x11);
                Sys.Desktop.KeyDown(0x46);
               Sys.Desktop.KeyUp(0x11);
                Sys.Desktop.KeyUp(0x46);
                        
       var invoice = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.invoice;
       Sys.HighlightObject(invoice);       
       ReportUtils.logStep_Screenshot(""); 
       invoice.Click();
       aqUtils.Delay(2000,Indicator.Text);
       
       var invoicehistory = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.invoicehis;
       Sys.HighlightObject(invoicehistory);
       invoicehistory.Click();
       aqUtils.Delay(2000,Indicator.Text);
       
       var invoicetable = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.invoicetable;
       Sys.HighlightObject(invoicetable);
       ReportUtils.logStep_Screenshot(""); 
       var row = invoicetable.getItemCount();
       var column = invoicetable.getColumnCount();
       
         for(var i=0;i<invoicetable.getItemCount();i++){   
//           if(invoicetable.getItem(i).getText(0).OleValue.toString().trim()==IvoiceNumber){       
                 if(invoicetable.getItem(i).getText(8).OleValue.toString().trim()=="Not Due"){  
//                    var getamount = invoicetable.getItem(i).getText(5).OleValue.toString().trim();
//                    ValidationUtils.verify(true,true,"Invoice Amount is : "+getamount);
//                    ExcelUtils.setExcelName(workBook,"Data Management", true);
//                    ExcelUtils.WriteExcelSheet("InvoiceAmount",EnvParams.Opco,"Data Management",getamount)
                    break;
                  }
                 else{
                       invoicetable.Keys("[Down]");
                 } 
//           }
//           else{
//               invoicetable.Keys("[Down]");
//           }                  
      
      }
      ReportUtils.logStep_Screenshot(""); 
      ValidationUtils.verify(true,true,"Payment status Need to Reconcile");
      
      var Home = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.overview;        
        Sys.HighlightObject(Home);
        Home.Click();
        
       aqUtils.Delay(3000,Indicator.Text);        
       var clientbalance = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.clientbalance;
       Clientbalance = clientbalance.getText();
       ReportUtils.logStep_Screenshot(); 
       ValidationUtils.verify(true,true,"Payment status Need to Reconcile");
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
  }

function gotoReceivable(){  
      
//      if((invoiceamount=="")||(invoiceamount==null)){
//      ExcelUtils.setExcelName(workBook, "Data Management", true);
//      invoiceamount = ReadExcelSheet("InvoiceAmount",EnvParams.Opco,"Data Management");
//      Log.Message(invoiceamount);
//      }

     var clientopen = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Registrations;
     clientopen.HoverMouse();
     ReportUtils.logStep_Screenshot("");
     clientopen.Click();
     aqUtils.Delay(2000,Indicator.Text);
     
     var newbutton = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.newbutton;
     Sys.HighlightObject(newbutton);
     newbutton.Click();
     aqUtils.Delay(3000,Indicator.Text);
     
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
     
     var company = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.company;
     company.Click();
     if(company!=""){
       company.setText(companyno);
       aqUtils.Delay(2000,Indicator.Text);
       ValidationUtils.verify(true,true,"Company is Entered");
     }
     else{
       ValidationUtils.verify(false,true,"Company is Needed to Create Single Invoice");
     }
     aqUtils.Delay(2000,Indicator.Text);
     
    var descrip = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.descrip;
     Sys.HighlightObject(descrip);     
     descrip.Click();
     if(Descip!=""){
      descrip.setText(Descip);
       aqUtils.Delay(2000,Indicator.Text);
       ValidationUtils.verify(true,true,"Description is Entered");
     }
     else{
       ValidationUtils.verify(false,true,"Description is Needed to Create Multiple Invoice");
     }
     
     var currenccy = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite2.currency;
      if(currency!=""){
       currenccy.Click();
       aqUtils.Delay(2000, Indicator.Text);
       WorkspaceUtils.DropDownList(currency,"Currency");
       aqUtils.Delay(2000, Indicator.Text); 
    } 
    else{
      ValidationUtils.verify(false,true,"Currency is Needed to Create Single Invoice"); 
    } 
     
    var amount = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite3.amount;
     Sys.HighlightObject(amount);
     amount.Click();    
      amount.setText("^a[BS]");
     if(Clientbalance!=""){       
       amount.setText(Clientbalance);
       ValidationUtils.verify(true,true,"Invoice Amount is Entered");
       aqUtils.Delay(2000,Indicator.Text);
     }
     else{
       ValidationUtils.verify(false,true,"Invoice Amount is Needed to Create Single Invoice"); 
     }
     
     var scroll= Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
     scroll.Click();   
     scroll.MouseWheel(-200);
      
     var client = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.Composite.client;
      if(clientnum!=""){
       client.Click();
          WorkspaceUtils.SearchByValuePicker(client,"Client",clientnum,"Client Number");
      } 
      else{ 
      ValidationUtils.verify(false,true,"Client Number is Exist for Single Invoice");
      } 
      
      scroll.MouseWheel(200);
      
      var showbutton = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget.Composite.McPlainCheckboxView.showButton;
      showbutton.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      showbutton.Click();
      ReportUtils.logStep("INFO", "Show Lines is Checked");

      var save = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.savee;
      save.Click();
      aqUtils.Delay(3000,Indicator.Text);
     
       var getjournal = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite.journal.getText();
       ValidationUtils.verify(true,true,"Journal Number is : "+getjournal);
       ExcelUtils.setExcelName(workBook,"Data Management", true);
       ExcelUtils.WriteExcelSheet("Journal No",EnvParams.Opco,"Data Management",getjournal)
       
       var artable = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table;
       Sys.HighlightObject(artable);       
       artable.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
       

       var recon = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.amount;
       recon.setText(invoiceamount);
       aqUtils.Delay(3000,Indicator.Text);
       recon.Keys("[Tab]");
       
       var tp = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.assettype;
         if(TP!=""){
       tp.Click();
       aqUtils.Delay(2000, Indicator.Text);
       WorkspaceUtils.DropDownList(TP,"Totally");
       aqUtils.Delay(2000, Indicator.Text); 
      } 
      else{
        ValidationUtils.verify(false,true,"TP is Needed to Create Single Invoice");  
      } 
    aqUtils.Delay(1000,Indicator.Text);
    var saveentry = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
     saveentry.Click();
      aqUtils.Delay(2000,Indicator.Text); 
       
       var released = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
       Sys.HighlightObject(released);
       released.Click();
       aqUtils.Delay(2000,Indicator.Text);
       
       var relesave = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 3);
       relesave.Click();          
       aqUtils.Delay(2000,Indicator.Text); 
      
      var clientpayment = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.clientpayment;
      Sys.HighlightObject(clientpayment);
      clientpayment.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      clientpayment.Click();
      aqUtils.Delay(2000,Indicator.Text);
      
      var tab = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
      Sys.HighlightObject(tab);
      
      var journalnum = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.journalnumber;
      journalnum.Click();
      journalnum.setText(getjournal);
      aqUtils.Delay(4000,Indicator.Text);     
      
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);     
      
      var submit = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.submitjournal;
      Sys.HighlightObject(submit);
      submit.Click();
      aqUtils.Delay(4000,Indicator.Text); 
      submit.HoverMouse();
      ReportUtils.logStep_Screenshot("");    
      ValidationUtils.verify(true,true,"Journal Number is Submitted");  
      
}  
  


function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}




 
 