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
var sheetName = "Reverse Invoice";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");
   
var STIME ="";
var companyno ="";
var Invoiceno = "";
var Invoicenew ="";
var Description ="";
var vendor ="";

  
function ReverseInvoice() {    
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
      goToAPTrans();
      GoToreverse();
}



function getDetails(){             
        sheetName = "Reverse Invoice";           
        ExcelUtils.setExcelName(workBook, sheetName, true);
        
        companyno = ExcelUtils.getRowDatas("company",EnvParams.Opco)
        if((companyno==null)||(companyno=="")){ 
        ValidationUtils.verify(false,true,"Company Number is needed to Reverse Invoice"); 
        } 
        Description = ExcelUtils.getRowDatas("Descrip",EnvParams.Opco)
        Log.Message(Description)
        if((Description==null)||(Description=="")){ 
        ValidationUtils.verify(false,true,"Description is needed to Reverse Invoice"); 
        }
        Invoicenew = ExcelUtils.getRowDatas("NewInvoiceNo",EnvParams.Opco)
        Log.Message(Invoicenew)
        if((Invoicenew==null)||(Invoicenew=="")){ 
        ValidationUtils.verify(false,true,"Vendor Number is needed to Reverse Invoice"); 
        }
        Invoiceno = ExcelUtils.getRowDatas("InvoiceNo",EnvParams.Opco)
        Log.Message(Invoiceno)
        if((Invoiceno=="")||(Invoiceno==null)){
          ExcelUtils.setExcelName(workBook, "Data Management", true);
          Invoiceno = ReadExcelSheet("InvoiceNo",EnvParams.Opco,"Data Management");
        }  
        if((Invoiceno==null)||(Invoiceno=="")){ 
        ValidationUtils.verify(false,true,"Invoice Number is needed to Reverse Invoice"); 
        }        
        
} 


function address(){
  var company = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.companynumberr.getText().OleValue.toString().trim();
    if(company!="Company"){
      ValidationUtils.verify(false,true,"Company field is missing in macanomy for Reverse Invoice");
    }
    else{
      ValidationUtils.verify(true,true,"Company field is available in Macanomy for Reverse Invoice"); 
    }
  
  var invoiceno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.invoiceno.getText().OleValue.toString().trim();
    if(invoiceno!="Invoice No."){
      ValidationUtils.verify(false,true,"Invoice Number field is missing in macanomy for Reverse Invoice");
    }
    else{
      ValidationUtils.verify(true,true,"Invoice field is available in Macanomy for Reverse Invoice"); 
    }
    
   var Descripton = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite2.descrip.getText().OleValue.toString().trim();
    if(Descripton!="Description"){
      ValidationUtils.verify(false,true,"Description field is available in Macanomy for Reverse Invoice");
    }
    else{
      ValidationUtils.verify(true,true,"Description field is missing in macanomy for Reverse Invoice");
    } 
    var vendor = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.Composite.Vendorfield.getText().OleValue.toString().trim();
    if(vendor!="Vendor"){
      ValidationUtils.verify(false,true,"Vendor field is missing in macanomy for Reverse Invoice");
    }
    else{
      ValidationUtils.verify(true,true,"Vendor field is available in Macanomy for Reverse Invoice");
    }
}


function goToAPTrans(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet0.Account_Payable.Exists()){
         ImageRepository.ImageSet0.Account_Payable.Click();// GL
         }
        else if(ImageRepository.ImageSet0.Account_Payable_1.Exists()){
        ImageRepository.ImageSet0.Account_Payable_1.Click();
        }
        else{
        ImageRepository.ImageSet0.Account_Payable_2.Click();
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
        Client_Managt.ClickItem("|AP Transactions");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|AP Transactions");
      }
    }
    aqUtils.Delay(3000,Indicator.Text);
  }

  
  
function GoToreverse(){
  
  var vendorinvoice = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.clientpayment;
  Sys.HighlightObject(vendorinvoice);
  aqUtils.Delay(1000,Indicator.Text);
  
  var invoiceallocation = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  invoiceallocation.Click();
  aqUtils.Delay(3000,Indicator.Text);
  
     ReportUtils.logStep_Screenshot("");
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
        aqUtils.Delay(2000,Indicator.Text);
  
  var addicon = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite3.Addiconbutton;
  addicon.Click();
  aqUtils.Delay(3000,Indicator.Text);
  
     
           
    address();     
    
    var companynum = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.companyfield;
    Sys.HighlightObject(companynum);
    if(companyno!=""){
      companynum.Click();
      WorkspaceUtils.SearchByValue(companynum,"Company",companyno,"Company No");
    }
    else{
      ValidationUtils.verify(false,true,"Company Number is needed to Create Writing Off Bad Debts");
    }   
    
    var invoicenew = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.invoicenumber;    
    if(Invoicenew!=""){
      invoicenew.Click();
        invoicenew.setText(Invoicenew);
        ValidationUtils.verify(true,true,"Invoice Number is Entered");
    }
    else{
        ValidationUtils.verify(false,true,"Invoice Number is Needed to Reverse Invoice");
    }    
    
    var Descrip = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite2.descripfield;
    if(Description!=""){
      Descrip.Click();
        Descrip.setText(Description);
        ValidationUtils.verify(true,true,"Description is Entered");
    }
    else{
        ValidationUtils.verify(false,true,"Description is Needed to Reverse Invoice");
    }  
    var scroll = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
    scroll.Click();
    scroll.MouseWheel(-200);

      
//    var vendornum = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget2.Composite.Composite.vendornum;
//    if(vendor!=""){
//      vendornum.Click();
//        WorkspaceUtils.SearchByValuePicker(vendornum,"Vendor",vendor);
//    }
//    else{
//        ValidationUtils.verify(false,true,"Vendor Number is Needed to Reverse Invoice");
//    } 
//    var amount = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite4.amountfield;
//    if(Amount!=""){
//      amount.Click();
//      amount.setText(Amount);
//         ValidationUtils.verify(true,true,"Amount is Entered");
//    }
//    else{
//        ValidationUtils.verify(false,true,"Amount is Needed to Reverse Invoice");
//    }    
//    ReportUtils.logStep_Screenshot("");
//    
    var copyinvoice = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite3.copyinvoice;
    if(Invoiceno!=""){
      copyinvoice.Click();
      WorkspaceUtils.SearchByValuePickerInvoice(copyinvoice,"Vendor Invoice",Invoiceno);
    }
     else{
       ValidationUtils.verify(false,true,"Invoice Number is needed to Reverse Invoice");
     }
    aqUtils.Delay(3000,Indicator.Text);    
    
    var reverse = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite.McPlainCheckboxView.ReverseCopyButton;
    var originalERate = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.McGroupWidget.Composite2.McPlainCheckboxView.OriginalERateButton;
           reverse.HoverMouse();
            ReportUtils.logStep_Screenshot("");
            reverse.Click();
            ValidationUtils.verify(true,true,"Reverse at Copying is Checked");
            ReportUtils.logStep("INFO", "Reverse at Copying is Checked");
            originalERate.HoverMouse();
            ReportUtils.logStep_Screenshot("");
            originalERate.Click();
            ValidationUtils.verify(true,true,"Original Rate is Checked");
            ReportUtils.logStep("INFO", "Original Rate is Checked");
    aqUtils.Delay(2000,Indicator.Text);
    var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.save;
    save.Click();
    aqUtils.Delay(2000,Indicator.Text);
    scroll.Click();
    scroll.MouseWheel(+400);
    aqUtils.Delay(3000,Indicator.Text);
    var invoicetype =NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite3.invoicetype.Click();
    var Type = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite3.invoicetype.getText();
    ValidationUtils.verify(true,true,"Invoice Type Changed as:"+Type);

     var action = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite3.attach;
    action.HoverMouse();
    action.Click();
    aqUtils.Delay(3000,Indicator.Text);
    Sys.Process("Maconomy").Refresh();
    var table = NameMapping.Sys.Maconomy.Window("#32768", "", 1);
    Sys.HighlightObject(table);
    for(var i=0;i<=5;i++){
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
    }
    aqUtils.Delay(2000,Indicator.Text);
    ReportUtils.logStep_Screenshot();
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    aqUtils.Delay(1000,Indicator.Text);
    var uploadlocal =  Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1); 
      uploadlocal.Keys(workBook);
          Sys.Desktop.KeyDown(0x0D);
          Sys.Desktop.KeyUp(0x0D); 
            aqUtils.Delay(3000, Indicator.Text);
            ValidationUtils.verify(true,true,"Document is Uploaded")
     action.Click();
    aqUtils.Delay(4000,Indicator.Text);
    Sys.Process("Maconomy").Refresh();
    var table = NameMapping.Sys.Maconomy.Window("#32768", "", 1);
    Sys.HighlightObject(table);  
      for(var i=0;i<=3;i++){
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
      }    
      aqUtils.Delay(3000,Indicator.Text);
    ReportUtils.logStep_Screenshot(); 
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    aqUtils.Delay(1000,Indicator.Text);
}


