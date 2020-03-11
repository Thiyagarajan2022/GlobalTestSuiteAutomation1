//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateCompanyVendor";
var STIME = "";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
STIME = WorkspaceUtils.StartTime();
var VendorName,Currency,attn,CpyTaxCode,Mail,phone,Taxderivation,Paymentmode,payterm,Annualsupplier,Supplier="";
var VendorNumber ="";


function CreateCompanyVendor(){ 
  TextUtils.writeLog("Create Company Vendor Started");
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateCompanyVendor";
Currency,VendorName ="";
Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME);
  getDetails();
  gotoMenu();
  gotoVendorSearch();
  globalVendor();
  NewglobalVendor();
  Policy();
  globalVendorTable();
  AttachDocument();
  Information();
  ApprvalInformation();
       for(var i=level;i<ApproveInfo.length;i++){
            level=i;
            WorkspaceUtils.closeMaconomy();
            aqUtils.Delay(10000, Indicator.Text);
            var temp = ApproveInfo[i].split("*");
            Restart.login(temp[2]);
            aqUtils.Delay(5000, Indicator.Text);
            todo(temp[3]);
            FinalApproveClient(temp[1],temp[2],i);
       }
  WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){ 
        Indicator.PushText("Reading Data from Excel");
        ExcelUtils.setExcelName(workBook, sheetName, true);
        
      Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
      if((Currency==null)||(Currency=="")){ 
      ValidationUtils.verify(false,true,"Currency is Needed to Create Company Vendor");
      }      
      Language = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
      if((Language==null)||(Language=="")){ 
      ValidationUtils.verify(false,true,"Language is Needed to Create Company Vendor");
      }
      attn = ExcelUtils.getRowDatas("Attn",EnvParams.Opco)
      if((attn==null)||(attn=="")){ 
      ValidationUtils.verify(false,true,"Attn. is Needed to Create Company Vendor");
      }
      Mail = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
      if((Mail==null)||(Mail=="")){ 
      ValidationUtils.verify(false,true,"E-mail is Needed to Create Company Vendor");
      }
      phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
      if((phone==null)||(phone=="")){ 
      ValidationUtils.verify(false,true,"Phone is Needed to Create Company Vendor");
      }
      CpyTaxCode = ExcelUtils.getRowDatas("CompanyTaxCode",EnvParams.Opco)
      if((CpyTaxCode==null)||(CpyTaxCode=="")){ 
      ValidationUtils.verify(false,true,"CompanyTaxCode is Needed to Create Company Vendor");
      }
      Taxderivation = ExcelUtils.getRowDatas("TaxDerivation",EnvParams.Opco)
      if((Taxderivation==null)||(Taxderivation=="")){ 
      ValidationUtils.verify(false,true,"Tax Derivation is Needed to Create Company Vendor");
      }
      Paymentmode = ExcelUtils.getRowDatas("PaymentMode",EnvParams.Opco)
      if((Paymentmode==null)||(Paymentmode=="")){ 
      ValidationUtils.verify(false,true,"Client Payment Mode is Needed to Create Company Vendor");
      }
      payterm = ExcelUtils.getRowDatas("PaymentTerms",EnvParams.Opco)
      if((payterm==null)||(payterm=="")){ 
      ValidationUtils.verify(false,true,"Payment Terms is Needed to Create Company Vendor");
      }   
      Supplier = ExcelUtils.getRowDatas("supplier",EnvParams.Opco)
      if((Supplier==null)||(Supplier=="")){ 
      ValidationUtils.verify(false,true,"Supplier is Needed to Create Global Vendor");
      }
      Annualsupplier = ExcelUtils.getRowDatas("annualsupplier",EnvParams.Opco)
      if((Annualsupplier==null)||(Annualsupplier=="")){ 
      ValidationUtils.verify(false,true,"Annual Supplier is Needed to Create Global Vendor");
      }
      VendorName = ExcelUtils.getRowDatas("Vendor Name",EnvParams.Opco)
       if((VendorName=="")||(VendorName==null)){
        ExcelUtils.setExcelName(workBook, "Data Management", true);
        VendorName = ReadExcelSheet("Vendor Name",EnvParams.Opco,"Data Management");
        Log.Message(VendorName)
        }
      if((VendorName==null)||(VendorName=="")){ 
      ValidationUtils.verify(false,true,"Vendor Name is Needed to Create Company Vendor");
      }
      
      Indicator.PushText("Playback");
}



function gotoMenu(){ 
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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

var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
aqUtils.Delay(3000, Indicator.Text);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|Vendor Management");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Vendor Management");
}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Vendor Management from Accounts Payable Menu");
}

function gotoVendorSearch(){ 
 var CompanyNumber = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  waitForObj(CompanyNumber);
  CompanyNumber.Click();
  var ExlArray = getExcelData_Company("Settling_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");

 var curr = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
//  aqUtils.Delay(2000, Indicator.Text);
    
 var Vendorname = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
 Vendorname.HoverMouse();
 Sys.HighlightObject(Vendorname); 
 Vendorname.setText(VendorName);
  
 var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 save.HoverMouse();
 Sys.HighlightObject(save);
  save.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  TextUtils.writeLog("Company Number, Vendor Number, Currency has entered and Saved in Vendor Search screen");
}

function globalVendor(){ 
  var Gblvendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  Gblvendor.HoverMouse();
  Sys.HighlightObject(Gblvendor);
    Gblvendor.Click();
//  aqUtils.Delay(3000, Indicator.Text);
  
   var active = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;  
   waitForObj(active);
  active.HoverMouse();
  Sys.HighlightObject(active);
  active.Click();
  aqUtils.Delay(3000, "Reading from Global Vendor table");
  var Newcompanyvendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
  Newcompanyvendor.HoverMouse();
  Sys.HighlightObject(Newcompanyvendor);
  Newcompanyvendor.Click();
  }
  
////=======================Vendor Creation=============////////
function NewglobalVendor(){ 
  
    var SettlingCompny = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
    waitForObj(SettlingCompny)   
    SettlingCompny.Click();
    var ExlArray = getExcelData_Company("Settling_Company",EnvParams.Opco)
    WorkspaceUtils.config_with_Maconomy_Validation(SettlingCompny,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");

   var Attn = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
   Attn.setText(attn);
 
    var Email = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
    Email.setText(Mail);
  
   var RemittanceEmail = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
   RemittanceEmail.setText(Mail);

    var Phone = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
    Phone.setText(phone)
    
    var companyTaxCode = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPopupPickerWidget;
    if(CpyTaxCode!=""){
    companyTaxCode.Click();
    WorkspaceUtils.DropDownList(CpyTaxCode,"Company Tax Code")
    }

    var taxDerivation = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McValuePickerWidget;
    taxDerivation.setText(Taxderivation);
//    if(Taxderivation!=""){
//    taxDerivation.Click();
//    WorkspaceUtils.SearchByValue(taxDerivation,"Local Specification 6",Taxderivation,"Name");
//    }      
    
    var paymentTerms = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McPopupPickerWidget;
    if(payterm!=""){
    paymentTerms.Click();
    WorkspaceUtils.DropDownList(payterm,"Payment Terms")
    }

    var paymentMode = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McValuePickerWidget;
     if(Paymentmode!=""){
      paymentMode.Click();
      WorkspaceUtils.SearchByValue(paymentMode,"Payment Mode",Paymentmode,"Name");
    }   
     
   var Next = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite2.Composite.Composite.Button;
   Next.HoverMouse();
   ReportUtils.logStep_Screenshot() ;
   Next.Click(); 
   
  }   

function Policy(){
   var scroll = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10   
     Sys.HighlightObject(scroll)    
      
      Sys.Desktop.KeyDown(0x12);
      Sys.Desktop.KeyDown(0x20);
      Sys.Desktop.KeyUp(0x12);
      Sys.Desktop.KeyUp(0x20);
      Sys.Desktop.KeyDown(0x58);
      Sys.Desktop.KeyUp(0x58);  
      aqUtils.Delay(1000, "Maximize the screen");
      
      var scroll = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10;
      scroll.MouseWheel(-200);
     
     var policy = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.McPopupPickerWidget;
     policy.Keys("yes");
     
     var nextpage = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite2.Composite.Composite.Button;
     waitForObj(nextpage);
     ReportUtils.logStep_Screenshot() ;
     nextpage.Click(); 
     
     var supplier = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.McTextWidget;
    
      Sys.Desktop.KeyDown(0x12);
      Sys.Desktop.KeyDown(0x20);
      Sys.Desktop.KeyUp(0x12);
      Sys.Desktop.KeyUp(0x20);
      Sys.Desktop.KeyDown(0x58);
      Sys.Desktop.KeyUp(0x58);  
      Sys.HighlightObject(supplier);
      Sys.HighlightObject(supplier);
      aqUtils.Delay(3000, Indicator.Text);
     
      supplier.Keys("yes");
     
     var PreferredSupplier = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.McPopupPickerWidget;     
    Sys.HighlightObject(PreferredSupplier)
      PreferredSupplier.Keys("yes");
     
     var newsupplier = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
     Sys.HighlightObject(newsupplier)
  newsupplier.Keys("yes");
     
     var duediligence = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
  Sys.HighlightObject(duediligence)   
  duediligence.Keys("yes");
     
     var servicerequired = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McPopupPickerWidget;     
    Sys.HighlightObject(servicerequired)  ;
   servicerequired.Keys("yes");    
   
    var abilitytodeliver = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.McTextWidget;
    Sys.HighlightObject(abilitytodeliver)  ;
     abilitytodeliver.Keys("yes"); 
    
     var agencyemployee = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPopupPickerWidget;
     Sys.HighlightObject(agencyemployee)  ;
    agencyemployee.Keys("yes");
     
     var impactrequest = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McTextWidget;
    Sys.HighlightObject(impactrequest) 
    impactrequest.Keys("yes");
     
    var suppliercurrency = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McTextWidget;
    Sys.HighlightObject(suppliercurrency) ;
    suppliercurrency.setText(Supplier)
    
     var annualsuppliercurrency = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.McTextWidget;
      annualsuppliercurrency.setText(Annualsupplier)
     aqUtils.Delay(1000, Indicator.Text);     
     
     var btnCreate = Aliases.Maconomy.CompanyVendor.Composite.Composite.Composite2.Composite.Button;
     Sys.HighlightObject(btnCreate);
      btnCreate.Click();           
      aqUtils.Delay(1000, Indicator.Text);
      
      
//       if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Vendors - Information")    
//        {
//        var button = Sys.Process("Maconomy").SWTObject("Shell", "Vendors - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//        var label = Sys.Process("Maconomy").SWTObject("Shell", "Vendors - Information").SWTObject("Label", "*").WndCaption;
//                   button.HoverMouse();
//               waitForObj(button);
//            Sys.HighlightObject(button);
//            button.HoverMouse();
//            button.Click();   
//                    
//         }       
//         
//    
//      var popup = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Management - Company Specific Vendor Information Card");
//      Sys.HighlightObject(popup);
//      var popupok = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Management - Company Specific Vendor Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      Sys.HighlightObject(popupok);
//      popupok.Click();
      
      
    var Label = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Management - Company Specific Vendor Information Card").SWTObject("Label", "*").getText();    
    ReportUtils.logStep("INFO","Label");
    var OK = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Management - Company Specific Vendor Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");    
    
    OK.Click();      
  }
  
  
  function globalVendorTable(){ 
    
    var companyvendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  companyvendor.HoverMouse();   
 Sys.HighlightObject(companyvendor);
        companyvendor.HoverMouse();
        companyvendor.HoverMouse();
        companyvendor.Click();
        
        var blocked = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
        blocked.HoverMouse();
        Sys.HighlightObject(blocked);
        blocked.HoverMouse();
        blocked.HoverMouse();
        blocked.Click();
         
      var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;      
      Sys.HighlightObject(table);
      var vendorname = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget3;
      vendorname.Click();
      vendorname.setText(VendorName)
      vendorname.HoverMouse();
      vendorname.HoverMouse();
      vendorname.HoverMouse();
       aqUtils.Delay(3000, "Reading Table Data");
       
      if(table.getItem(0).getText_2(1).OleValue.toString().trim()==VendorName){
      VendorName = table.getItem(0).getText_2(1).OleValue.toString().trim();
      table.HoverMouse(51, 60);
      ReportUtils.logStep_Screenshot();
      table.Click(51, 60);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }
      else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==VendorName){
      VendorName = table.getItem(1).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(51, 79);
      ReportUtils.logStep_Screenshot();  
      table.Click(51, 79);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }
      else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==VendorName){
      VendorName = table.getItem(2).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(51, 98);
      ReportUtils.logStep_Screenshot();
      table.Click(51, 98);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }
      else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==VendorName){
        VendorName = table.getItem(3).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(51, 117);
      ReportUtils.logStep_Screenshot();
      table.Click(51, 117);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }   
      
}


function AttachDocument(){ 
   var doc = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;   
  Sys.HighlightObject(doc);
  doc.HoverMouse();
  doc.HoverMouse();
  waitForObj(doc);
  doc.Click();
  var attchDocument = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  Sys.HighlightObject(attchDocument);
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  ReportUtils.logStep_Screenshot();
  attchDocument.Click();
  aqUtils.Delay(4000, "Attaching Document");
  var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);  
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Attaching Document");
}

function Information(){ 
  var info = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  info.HoverMouse();
  info.HoverMouse();
  info.HoverMouse();
  Sys.HighlightObject(info);
  info.HoverMouse();
  info.HoverMouse();
  info.Click();
  aqUtils.Delay(2000, "Playback");
  var submit = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;  
  Sys.HighlightObject(submit);
  submit.HoverMouse();
  submit.HoverMouse();
  submit.Click();
  aqUtils.Delay(2000, "Playback");
  
}


function ApprvalInformation(){ 

        VendorNumber = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget.getText();        
        Log.Message("Vendor Number :" + VendorNumber);
         aqUtils.Delay(3000, Indicator.Text);
         
         
       var VendorApprovalpane = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;       
        Sys.HighlightObject(VendorApprovalpane);
        VendorApprovalpane.HoverMouse();
        VendorApprovalpane.Click();
        if(ImageRepository.ImageSet0.Maximize.Exists()){
        ImageRepository.ImageSet0.Maximize.Click();
        }
        var ClientVendorApproval =  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;        
       Sys.HighlightObject(ClientVendorApproval);
        ClientVendorApproval.HoverMouse();
        ClientVendorApproval.Click();
           var ApproverTable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
           var y=0;
              for(var i=0;i<ApproverTable.getItemCount();i++){   
                 var approvers="";
                  if(ApproverTable.getItem(i).getText_2(3)!="Approved"){
                  approvers = EnvParams.Opco+"*"+VendorNumber+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(5).OleValue.toString().trim();
                  Log.Message("Approver level :" +i+ ": " +approvers);
                  Approve_Level[y] = approvers;
                  Log.Message(Approve_Level[y])
                  y++;
                  }
              }
           TextUtils.writeLog("Finding approvers for Created Global Vendor");
        var closeCAList = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl;
      Sys.HighlightObject(closeCAList);
        closeCAList.HoverMouse();
        closeCAList.Click();
        ImageRepository.ImageSet0.Forward.Click();
        
        CredentialLogin();
        var OpCo2 = ApproveInfo[0].split("*");
        ExcelUtils.setExcelName(workBook, "Server Details", true);
        var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);

        sheetName = "CreateCompanyVendor";
        if(OpCo2[2]==Project_manager){
        level = 1;
        var Approve = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
        Sys.HighlightObject(Approve)
        if(Approve.isEnabled()){ 
        Approve.HoverMouse();
        ReportUtils.logStep_Screenshot();
        Approve.Click();
        aqUtils.Delay(8000, "Waiting for Approve");;
        ValidationUtils.verify(true,true,"Compnay Vendor is Approved by "+Project_manager)
        TextUtils.writeLog("Levels 0 has  Approved the Global Vendor");
        }
        }
}

function CredentialLogin(){ 
  var AppvLevl = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }
//WorkspaceUtils.closeAllWorkspaces();
}



function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
 
    var linestatus = false;
    if(!linestatus) 
    if((Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite).isVisible())
    {
    var toDo = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
    linestatus = true;
    }
     if(!linestatus) 
    if((Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite2).isVisible())
    {   
    var toDo = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
     linestatus = true;
    }
       if(!linestatus) 
    if((Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite4).isVisible())
    {   
    var toDo = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite4.PTabFolder.TabFolderPanel.TabControl;
     linestatus = true;
    }
    
    
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);  

  var linestatus = false;
    if(!linestatus) 
    if((Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite2).isVisible())
    {
    var refresh = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
    linestatus = true;
    }
     if(!linestatus) 
    if((Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite4).isVisible())
    {   
    var refresh = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite4.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;  
    linestatus = true;
    }
     if(!linestatus) 
    if((Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite3).isVisible())
    {   
    var refresh = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
    linestatus = true;
    }
    

  refresh.Click();
  aqUtils.Delay(15000, Indicator.Text);  

    var linestatus = false;
    if(!linestatus) 
    if((Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite2).isVisible())
    {
    var Client_Managt = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
    linestatus = true;
    }
     if(!linestatus) 
    if((Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite4).isVisible())
    {   
    var Client_Managt = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite4.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
    linestatus = true;
    }    
      if(!linestatus) 
    if((Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite3).isVisible())
    {   
    var Client_Managt = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
    linestatus = true;
    }


  var listPass = true;
      if(lvl==2)
        for(var j=0;j<Client_Managt.getItemCount();j++){
          var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
          var temp1 = temp.split("(");
          if((temp.indexOf("Approve Company Vendor by Type (")!=-1)&&(temp1.length==2)){ 
            Client_Managt.ClickItem("|"+temp);   
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp);  
            TextUtils.writeLog("Entering into Approve Company Vendor by Type from To-Dos List");
            listPass = false; 
          }
      }
      if(lvl==3)
      for(var j=0;j<Client_Managt.getItemCount();j++){ 
          var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
          var temp1 = temp.split("(");
        if((temp.indexOf("Approve Company Vendor by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
            Client_Managt.ClickItem("|"+temp);    
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp); 
            TextUtils.writeLog("Entering into Approve Company Vendor by Type (Substitute) from To-Dos List");
            var listPass = true;   
         }
      }  
  if(listPass){
    if(lvl==2)
          for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
              if((temp.indexOf("Approve Company Vendor (")!=-1)&&(temp1.length==2)){ 
                  Client_Managt.ClickItem("|"+temp);   
                  ReportUtils.logStep_Screenshot(); 
                  Client_Managt.DblClickItem("|"+temp);  
                  TextUtils.writeLog("Entering into Approve Company Vendor from To-Dos List");
                  listPass = false; 
                }
           }
    if(lvl==3)
        for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
          if((temp.indexOf("Approve Company Vendor (Substitute) (")!=-1)&&(temp1.length==3)){ 
              Client_Managt.ClickItem("|"+temp);    
              ReportUtils.logStep_Screenshot(); 
              Client_Managt.DblClickItem("|"+temp); 
              TextUtils.writeLog("Entering into Approve Company Vendor (Substitute) from To-Dos List");
              var listPass = true;   
            }
        } 
  }
}


function FinalApproveClient(VendorNum,Apvr,lvl){ 
   var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder;
    waitForObj(table);
    Sys.HighlightObject(table);
      if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Visible){
      }
      else{
      var showFilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;      
      waitForObj(table);
      Sys.HighlightObject(showFilter);
      showFilter.HoverMouse();
      showFilter.HoverMouse();
      showFilter.HoverMouse();
      showFilter.Click();
      }
    var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
    var cell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
    cell.Click();
    cell.Keys("[Tab]");
    var firstCell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
    waitForObj(firstCell);
    Sys.HighlightObject(firstCell);
    firstCell.HoverMouse();
    firstCell.HoverMouse();
    firstCell.setText(VendorNum);
    aqUtils.Delay(3000, "Reading Data in table");;
    var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
    waitForObj(closefilter);
    Sys.HighlightObject(closefilter);
    closefilter.HoverMouse();
    closefilter.HoverMouse(); 
    closefilter.HoverMouse();
    closefilter.HoverMouse(); 
      var flag=false;
      for(var v=0;v<table.getItemCount();v++){ 
        if(table.getItem(v).getText_2(1).OleValue.toString().trim()==VendorNum){ 
          flag=true;    
          break;
        }
        else{ 
          table.Keys("[Down]");
        }
      }

  ValidationUtils.verify(flag,true,"Created Company Vendor is available in Approval List");
  TextUtils.writeLog("Created Company Vendor is available in Approval List");
      if(flag){ 
      closefilter.HoverMouse();
      ReportUtils.logStep_Screenshot();
      closefilter.Click();
      aqUtils.Delay(5000, Indicator.Text);

      var Approve = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;      
      Sys.HighlightObject(Approve)
      if(Approve.isEnabled()){ 
      Approve.HoverMouse();
      ReportUtils.logStep_Screenshot();
      Approve.Click();      
      aqUtils.Delay(8000, "Waiting To Approve");;
      ValidationUtils.verify(true,true,"Company Vendor is Approved by "+Apvr)
      aqUtils.Delay(8000, Indicator.Text);
      TextUtils.writeLog("Company Vendor is Approved by "+Apvr);
      if(Approve_Level.length==lvl+1){
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Vendor by Type").SWTObject("Label", "*");
//      Log.Message(label.getText());
//      var lab = label.getText().OleValue.toString().trim();
//      var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Vendor by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      Ok.HoverMouse(); 
//      ReportUtils.logStep_Screenshot();
//      Ok.Click(); 
//      aqUtils.Delay(8000, Indicator.Text); ;
//       for(var j=0;j<12;j++){ 
//      if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Approve Company Vendor by Type"){ 
//      var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Vendor by Type").SWTObject("Label", "*");
//      Log.Message(label.getText());
//      var lab = label.getText().OleValue.toString().trim();
//      var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Vendor by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      Ok.HoverMouse(); 
//      ReportUtils.logStep_Screenshot();
//      Ok.Click(); 
//      aqUtils.Delay(8000, Indicator.Text); 
//      }
//      }
       TextUtils.writeLog("Vendor Number :"+ VendorNum); 
       var VendorApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;        
       Sys.HighlightObject(VendorApproval);
       VendorApproval.HoverMouse();
       VendorApproval.Click();
      // }
       if(ImageRepository.ImageSet.Maximize.Exists()){
      ImageRepository.ImageSet.Maximize.Click();
      }
       var VendorApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2       
       Sys.HighlightObject(VendorApproval);
       VendorApproval.HoverMouse();
       VendorApproval.Click();
         var ApproverTable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
        Sys.HighlightObject(ApproverTable);
        ReportUtils.logStep_Screenshot();
        var closeApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl;
        Sys.HighlightObject(closeApproval);
       closeApproval.HoverMouse();
       closeApproval.Click();
       ImageRepository.ImageSet.Forward.Click();       
       var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
        menuBar.Click();
      }
        ValidationUtils.verify(true,true,"Company Vendor is Approved by "+Apvr)  
      }
      }
}  








function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
var temp ="";
while (!DDT.CurrentDriver.EOF()) {
if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
try{
temp = temp+xlDriver.Value(column).toString().trim();
}
catch(e){
temp = "";
}
break;
}
xlDriver.Next();
}
     
if(temp.indexOf("*")!=-1){
var excelData =  temp.split("*");
}else if(temp.length>0){ 
excelData[0] = temp;
}
     
DDT.CloseDriver(xlDriver.Name);
for(var i=0;i<excelData.length;i++)
return excelData;
}





function GVData_1_Address(){ 
  aqUtils.Delay(4000, Indicator.Text);
Sys.Process("Maconomy").Refresh();
var country_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget).getText().OleValue.toString().trim()
if(country_1!="Country")
ValidationUtils.verify(false,true,"Country field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Country field is available in Maconomy for Vendor Creation");
var taxcode_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget).getText().OleValue.toString().trim()
if(taxcode_1!="Tax No.")
ValidationUtils.verify(false,true,"Tax No. field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Tax No. field is available in Maconomy for Vendor Creation");
var companyReg_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget).getText().OleValue.toString().trim()
if(companyReg_1!="Company Reg. No.")
ValidationUtils.verify(false,true,"Company Reg. No. field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Company Reg. No. field is available in Maconomy for Vendor Creation");
var currency_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget).getText().OleValue.toString().trim()
if(currency_1!="Currency")
ValidationUtils.verify(false,true,"Currency field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Currency field is available in Maconomy for Vendor Creation");
var clientgrp_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget).getText().OleValue.toString().trim()
if(clientgrp_1!="Vendor Group")
ValidationUtils.verify(false,true,"Client Group field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Client Group field is available in Maconomy for Vendor Creation");
var controlAct_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget).getText().OleValue.toString().trim()
if(controlAct_1!="Control Account")
ValidationUtils.verify(false,true,"Control Account field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Control Account field is available in Maconomy for Vendor Creation");
var bfc_1 = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McTextWidget).getText().OleValue.toString().trim()
if(bfc_1!="Counter Party BFC")
ValidationUtils.verify(false,true,"Counter Party BFC field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Counter Party BFC field is available in Maconomy for Vendor Creation");
var BankAccountName = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McTextWidget).getText().OleValue.toString().trim()
if(BankAccountName!="Bank Account Name")
ValidationUtils.verify(false,true,"Bank Account Name field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"Bank Account Name field is available in Maconomy for Vendor Creation");
var IBAN = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McTextWidget).getText().OleValue.toString().trim()
if(IBAN!="IBAN")
ValidationUtils.verify(false,true,"IBAN field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"IBAN field is available in Maconomy for Vendor Creation");

var SWIFT = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.McTextWidget).getText().OleValue.toString().trim()
if(SWIFT!="SWIFT")
ValidationUtils.verify(false,true,"SWIFT field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"SWIFT field is available in Maconomy for Vendor Creation");

var BankAcctNo = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.McTextWidget).getText().OleValue.toString().trim()
if(BankAcctNo!="Bank Acct. No.")
ValidationUtils.verify(false,true,"BankAcctNo field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"BankAcctNo field is available in Maconomy for Vendor Creation");

var SortCode = (Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.McTextWidget).getText().OleValue.toString().trim()
if(SortCode!="Sort Code / ABA No.")
ValidationUtils.verify(false,true,"SortCode field is missing in Maconomy for Vendor Creation");
else
ValidationUtils.verify(true,true,"SortCode field is available in Maconomy for Vendor Creation");
}


 