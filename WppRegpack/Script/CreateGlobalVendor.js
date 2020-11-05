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
var sheetName = "CreateGlobalVendor";
var STIME = "";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
STIME = WorkspaceUtils.StartTime();
var VendorName,strt1,post,City,country,taxcode,companyReg,Currency,VendorGroup,controlAct,BankAccountName,bfc,Iban,Swift,
BankActNo,Sortcode,company,attn,CpyTaxCode,Mail,phone,Taxderivation,Paymentmode,payterm,Annualsupplier,
Supplier,Method,Section,TDSAplicable,GSTVendor,StateCode,Vendortype,SII_Tax ="";
var VendorNumber ="";
var languagee="";
var Language = "";
function CreateGlobalVendor(){ 
  TextUtils.writeLog("Create Gloabl Vendor Started");
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);

var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateGlobalVendor";
Currency,VendorName ="";

STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME);

 try{
  getDetails();
  gotoMenu();
  gotoVendorSearch();
  globalVendor();
  NewglobalVendor();
  VendorScreen();
  Policy();
  globalVendorTable();
   if(EnvParams.Country.toUpperCase()=="INDIA"){
    Runner.CallMethod("IND_GlobalVendor.IndiaSpecific",Vendortype,StateCode,GSTVendor,TDSAplicable,Section,Method);
   }
   if(EnvParams.Country.toUpperCase()=="SPAIN"){
  Runner.CallMethod("SPA_CreateGlobalVendor.spainSpecific",SII_Tax);
  }
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
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
              
            }
            var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
       }
  }
    catch(err){
      Log.Message(err);
    }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){ 
        Indicator.PushText("Reading Data from Excel");
        ExcelUtils.setExcelName(workBook, sheetName, true);
        Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
        Log.Message(Currency)
        if((Currency==null)||(Currency=="")){ 
        ValidationUtils.verify(false,true,"Currency is Needed to Create Global Vendor");
        }
        
      VendorName = ExcelUtils.getRowDatas("Vendor Name",EnvParams.Opco)
      Log.Message(VendorName)
      if((VendorName==null)||(VendorName=="")){ 
      ValidationUtils.verify(false,true,"Vendor Name is Needed to Create Global Vendor");
      }
      strt1 = ExcelUtils.getRowDatas("Street 1",EnvParams.Opco)
      Log.Message(strt1)
      if((strt1==null)||(strt1=="")){ 
      ValidationUtils.verify(false,true,"Street 1 is Needed to Create Global Vendor");
      }      
      post = ExcelUtils.getRowDatas("Post",EnvParams.Opco)
      Log.Message(post)
      if((post==null)||(post=="")){ 
      ValidationUtils.verify(false,true,"Post is Needed to Create Global Vendor");
      }
      City = ExcelUtils.getRowDatas("City",EnvParams.Opco)
      Log.Message(City)
      if((City==null)||(City=="")){ 
      ValidationUtils.verify(false,true,"City is Needed to Create Global Vendor");
      }
      country = ExcelUtils.getRowDatas("Country",EnvParams.Opco)
      Log.Message(country)
      if((country==null)||(country=="")){ 
      ValidationUtils.verify(false,true,"Country is Needed to Create Global Vendor");
      }      
      taxcode = ExcelUtils.getRowDatas("Tax No",EnvParams.Opco)
      Log.Message(taxcode)
      if((taxcode==null)||(taxcode=="")){ 
      ValidationUtils.verify(false,true,"Tax No is Needed to Create Global Vendor");
      }
      companyReg = ExcelUtils.getRowDatas("CompRegNo",EnvParams.Opco)
      Log.Message(companyReg)
      if((companyReg==null)||(companyReg=="")){ 
      ValidationUtils.verify(false,true,"CompRegNo is Needed to Create Global Vendor");
      }   
      VendorGroup = ExcelUtils.getRowDatas("Vendor Group",EnvParams.Opco)
      Log.Message(VendorGroup)
      if((VendorGroup==null)||(VendorGroup=="")){ 
      ValidationUtils.verify(false,true,"Vendor Group is Needed to Create Global Vendor");
      }
      controlAct = ExcelUtils.getRowDatas("Control Account",EnvParams.Opco)
      Log.Message(controlAct)
      if((controlAct==null)||(controlAct=="")){ 
      ValidationUtils.verify(false,true,"Control Account is Needed to Create Global Vendor");
      }
      bfc = ExcelUtils.getRowDatas("BFC",EnvParams.Opco)
      Log.Message(bfc)
      if((bfc==null)||(bfc=="")){ 
      ValidationUtils.verify(false,true,"Counter Party BFC is Needed to Create Global Vendor");
      }   
      BankAccountName = ExcelUtils.getRowDatas("Bank Account Name",EnvParams.Opco)
      Log.Message(BankAccountName)
      if((BankAccountName==null)||(BankAccountName=="")){ 
      ValidationUtils.verify(false,true,"Bank Account Name is Needed to Create Global Vendor");
      }      
      Iban = ExcelUtils.getRowDatas("IBAN",EnvParams.Opco)
      Log.Message(Iban)
      if((Iban==null)||(Iban=="")){ 
      ValidationUtils.verify(false,true,"IBAN is Needed to Create Global Vendor");
      }      
      Swift = ExcelUtils.getRowDatas("SWIFT",EnvParams.Opco)
      Log.Message(Swift)
      if((Swift==null)||(Swift=="")){ 
      ValidationUtils.verify(false,true,"SWIFT is Needed to Create Global Vendor");
      }      
      BankActNo = ExcelUtils.getRowDatas("BankAcctNo",EnvParams.Opco)
      Log.Message(BankActNo)
      if((BankActNo==null)||(BankActNo=="")){ 
      ValidationUtils.verify(false,true,"BankAcctNo is Needed to Create Global Vendor");
      }      
      Sortcode = ExcelUtils.getRowDatas("SortCode",EnvParams.Opco)
      Log.Message(Sortcode)
      if((Sortcode==null)||(Sortcode=="")){ 
      ValidationUtils.verify(false,true,"SortCode is Needed to Create Global Vendor");
      }      
      languagee = ExcelUtils.getRowDatas("language",EnvParams.Opco)
      Log.Message(languagee)
      if((languagee==null)||(languagee==""))
       languagee = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
      if((languagee==null)||(languagee=="")){ 
      ValidationUtils.verify(false,true,"Language is Needed to Create Global Vendor");
      }
      attn = ExcelUtils.getRowDatas("Attn",EnvParams.Opco)
      if((attn==null)||(attn=="")){ 
      ValidationUtils.verify(false,true,"Attn. is Needed to Create Global Vendor");
      }
      Mail = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
      if((Mail==null)||(Mail=="")){ 
      ValidationUtils.verify(false,true,"E-mail is Needed to Create Global Vendor");
      }
      phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
      if((phone==null)||(phone=="")){ 
      ValidationUtils.verify(false,true,"Phone is Needed to Create Global Vendor");
      }
      CpyTaxCode = ExcelUtils.getRowDatas("CompanyTaxCode",EnvParams.Opco)
      if((CpyTaxCode==null)||(CpyTaxCode=="")){ 
      ValidationUtils.verify(false,true,"CompanyTaxCode is Needed to Create Global Vendor");
      }
      Taxderivation = ExcelUtils.getRowDatas("TaxDerivation",EnvParams.Opco)
      if((Taxderivation==null)||(Taxderivation=="")){ 
      ValidationUtils.verify(false,true,"Tax Derivation is Needed to Create Global Vendor");
      }
      Paymentmode = ExcelUtils.getRowDatas("PaymentMode",EnvParams.Opco)
      if((Paymentmode==null)||(Paymentmode=="")){ 
      ValidationUtils.verify(false,true,"Client Payment Mode is Needed to Create Global Vendor");
      }
      payterm = ExcelUtils.getRowDatas("PaymentTerms",EnvParams.Opco)
      if((payterm==null)||(payterm=="")){ 
      ValidationUtils.verify(false,true,"Payment Terms is Needed to Create Global Vendor");
      }
      Supplier = ExcelUtils.getRowDatas("supplier",EnvParams.Opco)
      Log.Message(Supplier)
      if((Supplier==null)||(Supplier=="")){ 
      ValidationUtils.verify(false,true,"Supplier is Needed to Create Global Vendor");
      }
      Annualsupplier = ExcelUtils.getRowDatas("annualsupplier",EnvParams.Opco)
      if((Annualsupplier==null)||(Annualsupplier=="")){ 
      ValidationUtils.verify(false,true,"Annual Supplier is Needed to Create Global Vendor");
      }
      
if(EnvParams.Country.toUpperCase()=="INDIA"){    
      Vendortype = ExcelUtils.getRowDatas("Vendortype",EnvParams.Opco)
      Log.Message(Vendortype)
      if((Vendortype==null)||(Vendortype=="")){ 
      ValidationUtils.verify(false,true,"Vendor Type is Needed to Create Global Vendor");
      }
      
      StateCode = ExcelUtils.getRowDatas("Statecode",EnvParams.Opco)
      Log.Message(StateCode)
      if((StateCode==null)||(StateCode=="")){ 
      ValidationUtils.verify(false,true,"State Code is Needed to Create Global Vendor");
      }
      GSTVendor = ExcelUtils.getRowDatas("GST Vendor Type",EnvParams.Opco)
      Log.Message(GSTVendor)
      if((GSTVendor==null)||(GSTVendor=="")){ 
      ValidationUtils.verify(false,true,"GST Vendor Type is Needed to Create Global Vendor");
      }
      TDSAplicable = ExcelUtils.getRowDatas("TDSApplicable",EnvParams.Opco)
      Log.Message(TDSAplicable)
      if((TDSAplicable==null)||(TDSAplicable=="")){ 
      ValidationUtils.verify(false,true,"TDS Applicable is Needed to Create Global Vendor");
      }
      Section = ExcelUtils.getRowDatas("TDS Section",EnvParams.Opco)
      Log.Message(Section)
      if((Section==null)||(Section=="")){ 
      ValidationUtils.verify(false,true,"TDS Section is Needed to Create Global Vendor");
      }

      Method = ExcelUtils.getRowDatas("WHMethod",EnvParams.Opco)
      Log.Message(Method)
      if((Method==null)||(Method=="")){ 
      ValidationUtils.verify(false,true,"TDS Section is Needed to Create Global Vendor");
      }
}
      
      if(EnvParams.Country.toUpperCase()=="SPAIN"){
      SII_Tax = ExcelUtils.getRowDatas("SII Tax Group",EnvParams.Opco)
      Log.Message(SII_Tax)
      if((SII_Tax==null)||(SII_Tax=="")){ 
      ValidationUtils.verify(false,true,"SII Tax Group is Needed to Create a Client");
      }

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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Management").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Management").OleValue.toString().trim());

}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Vendor Management from Accounts Payable Menu");
}

function gotoVendorSearch(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

 var CompanyNumber = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  waitForObj(CompanyNumber);
  CompanyNumber.Click();
  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");

 var curr = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
    
 var Vendorname = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
 Vendorname.HoverMouse();
 Sys.HighlightObject(Vendorname); 
 Vendorname.setText(VendorName.toString().trim()+" "+STIME);
  
 var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 save.Click();
  aqUtils.Delay(5000, Indicator.Text);
  TextUtils.writeLog("Company Number, Vendor Number, Currency has entered and Saved in Vendor Search screen");
}

function globalVendor(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var Gblvendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  Gblvendor.Click();
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var Newvendor = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Newvendor.Click();
  aqUtils.Delay(9000, Indicator.Text); 
  }
  
////=======================Vendor Creation=============////////
function NewglobalVendor(){ 
  
  var vendor_details1 = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.McTextWidget;  
  Sys.HighlightObject(vendor_details1);
  waitForObj(vendor_details1)
  vendor_details1.Click()
  vendor_details1.setText(VendorName+" "+STIME);
  VendorName = VendorName+" "+STIME;
//  VendorName = VendorName.toString().trim()+" "+STIME; 
  
  var vendor_details2 = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.McTextWidget;
  vendor_details2.setText(strt1);
  
  var Post = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite15.Composite.McValuePickerWidget;
  Post.setText(post);
  
  var city = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite15.Composite.McValuePickerWidget2;
  city.setText(City);
  
  var Country = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  if(country!=""){
  Country.Click();
  aqUtils.Delay(1000,"Loading dropdown values");
  WorkspaceUtils.DropDownList(country,"Country")
  }

  var tax = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget2;
  tax.setText(taxcode); 
  
  var regno = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget2;
  regno.setText(companyReg);  
  
  var currency = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
  waitForObj(currency); 
 if(Currency!=""){
  currency.Click();
  Log.Message(Currency)
  currency.Keys(Currency)
  aqUtils.Delay(5000,"Plackback")
//  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  
  var vgroup = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McPopupPickerWidget;
  if(VendorGroup!=""){
  vgroup.Click();
  WorkspaceUtils.DropDownList(VendorGroup,"Vendor Group")
  }
  
  var controlacc = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPopupPickerWidget;
  if(controlAct!=""){
  controlacc.Click();
  WorkspaceUtils.DropDownList(controlAct,"Control Account")
  }

  var CounterPartyBFC = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McValuePickerWidget;
  if(bfc!=""){
    CounterPartyBFC.Click();
    WorkspaceUtils.SearchByValue(CounterPartyBFC,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Counter Party BFC").OleValue.toString().trim(),bfc,"Counter Party BFC");
  }
  
  var bname = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McTextWidget2;
  bname.setText(BankAccountName);  
  
  var IBAN = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McTextWidget2
  IBAN.setText(Iban);  

  var swift = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.McTextWidget2;
  swift.setText(Swift); 
  
  var bankno = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.McTextWidget2;
  bankno.setText(BankActNo);  
  
  var sortcode = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.McTextWidget2;
  sortcode.setText(Sortcode);  
  
   var Next = Aliases.Maconomy.Group8.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
   Next.HoverMouse();
   ReportUtils.logStep_Screenshot() ;
   Next.Click(); 
  }
   
function VendorScreen(){
   var Compny = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.McValuePickerWidget;   
   waitForObj(Compny)   
    Compny.Click();
  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(Compny,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");

   var Attn = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.McTextWidget;
      Attn.setText(attn);
 
    var Email = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite16.McTextWidget;
    Email.setText(Mail);
 
    var RemittanceEmail = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite17.McTextWidget;
   RemittanceEmail.setText(Mail);
 
    var language = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite18.McPopupPickerWidget;
    if(languagee!=""){
    language.Click();
    WorkspaceUtils.DropDownList(languagee,"Language")
    }

    var Phone = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite15.McTextWidget;
    Phone.setText(phone)

    var companyTaxCode = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
   if(CpyTaxCode!=""){
    companyTaxCode.Click();
    WorkspaceUtils.DropDownList(CpyTaxCode,"Company Tax Code")
    }

    var taxDerivation = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
    if(Taxderivation!=""){
    taxDerivation.Click();
    taxDerivation.setText(Taxderivation)
//    WorkspaceUtils.SearchByValue(taxDerivation,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Local Specification 6").OleValue.toString().trim(),Taxderivation,"Name");
  }  
    
    Delay(5000)
   var paymentTerms = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
   if(payterm!=""){
    paymentTerms.Click();
    WorkspaceUtils.DropDownList(payterm,"Payment Terms")
    }

    var paymentMode = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McValuePickerWidget;
     if(Paymentmode!=""){
      paymentMode.Click();
      WorkspaceUtils.SearchByValue(paymentMode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Mode").OleValue.toString().trim(),Paymentmode,"Name");
    }   
     
     var next = Aliases.Maconomy.Group8.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
      next.HoverMouse();
     ReportUtils.logStep_Screenshot() ;
     next.Click(); 
    }
    
function Policy(){
   var scroll = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10;
     Sys.HighlightObject(scroll)    
      
      Sys.Desktop.KeyDown(0x12);
      Sys.Desktop.KeyDown(0x20);
      Sys.Desktop.KeyUp(0x12);
      Sys.Desktop.KeyUp(0x20);
      Sys.Desktop.KeyDown(0x58);
      Sys.Desktop.KeyUp(0x58);  
      aqUtils.Delay(1000, "Maximize the screen");
      
      
      var scroll = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10;
      scroll.Click();
      scroll.MouseWheel(-200);
      aqUtils.Delay(5000,"Plackback");
     var policy = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite19.McPopupPickerWidget;
     Sys.HighlightObject(policy)
     policy.Click();
     WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",policy)
//      policy.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var nextpage = Aliases.Maconomy.Group8.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
     waitForObj(nextpage);
     ReportUtils.logStep_Screenshot() ;
     nextpage.Click(); 
     
       var supplier = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.McTextWidget;
    
      Sys.Desktop.KeyDown(0x12);
      Sys.Desktop.KeyDown(0x20);
      Sys.Desktop.KeyUp(0x12);
      Sys.Desktop.KeyUp(0x20);
      Sys.Desktop.KeyDown(0x58);
      Sys.Desktop.KeyUp(0x58);  
      Sys.HighlightObject(supplier);
      Sys.HighlightObject(supplier);
//      aqUtils.Delay(3000, Indicator.Text);
     
      supplier.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var PreferredSupplier = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite16.McPopupPickerWidget;
    Sys.HighlightObject(PreferredSupplier)
    PreferredSupplier.Click();
     WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",PreferredSupplier)
//      PreferredSupplier.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var newsupplier = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite17.McTextWidget;
     Sys.HighlightObject(newsupplier)
  newsupplier.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var duediligence = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  Sys.HighlightObject(duediligence)   
  duediligence.Click();
     WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",duediligence)
//  duediligence.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var servicerequired = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McPopupPickerWidget;
    Sys.HighlightObject(servicerequired)  ;
    servicerequired.Click();
     WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",servicerequired)
//   servicerequired.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var deliver = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget2;
    Sys.HighlightObject(deliver) 
    deliver.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  
     var agencyemployee = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPopupPickerWidget;
     agencyemployee.Click();
     WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",agencyemployee)
//     agencyemployee.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
    
     var impactrequest = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McTextWidget2;
    Sys.HighlightObject(impactrequest) 
    impactrequest.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
      
     var suppliercurrency = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McTextWidget2;
    Sys.HighlightObject(suppliercurrency) ;
    suppliercurrency.setText(Supplier)
     
     var annualsuppliercurrency = Aliases.Maconomy.Group8.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.McTextWidget2;
      annualsuppliercurrency.setText(Annualsupplier)
     aqUtils.Delay(2000, Indicator.Text);
     var btnCreate = Aliases.Maconomy.Group8.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
     Sys.HighlightObject(btnCreate);
      btnCreate.Click();     
     
//     var btnCreate = Aliases.Maconomy.Group8.Composite.Composite.Composite2.Composite.Button;
//     if(btnCreate.isEnabled()){
//         Sys.HighlightObject(btnCreate)
//          btnCreate.HoverMouse();
//        ReportUtils.logStep_Screenshot("");
//          btnCreate.Click();
//          TextUtils.writeLog("Gloabl Vendor is CREATED");
//        ValidationUtils.verify(true,true,"Vendor is CREATED");
//      }
//      else{
//        var cancel = Aliases.Maconomy.Group8.Composite.Composite.Composite2.Composite.Button2;
//         Sys.HighlightObject(cancel)
//        cancel.HoverMouse();
//        ReportUtils.logStep_Screenshot("");
//        cancel.Click();
//        ValidationUtils.verify(false,true,"Vendor is Not CREATED");
//      }   
      
      aqUtils.Delay(3000, Indicator.Text);
      
    var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Management - Vendor Information Card").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
    ReportUtils.logStep("INFO","Label");
    var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Management - Vendor Information Card").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
    OK.Click();  
  }
  
  
  function globalVendorTable(){ 
        var blocked = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blocked").OleValue.toString().trim());
        Sys.HighlightObject(blocked);
        blocked.HoverMouse();
        blocked.HoverMouse();
        blocked.Click();
        blocked.HoverMouse();
        blocked.HoverMouse();
        blocked.HoverMouse();
         
      var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
      Sys.HighlightObject(table);
      aqUtils.Delay(8000, Indicator.Text);
      var vendorname = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
      vendorname.Click();
      vendorname.setText(VendorName)
      vendorname.HoverMouse();
      vendorname.HoverMouse();
      vendorname.HoverMouse();
       aqUtils.Delay(3000, "Reading Table Data");
       
      if(table.getItem(0).getText_2(1).OleValue.toString().trim()==VendorName){
      VendorNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(49, 52);
      ReportUtils.logStep_Screenshot();
      table.Click(49, 52);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }
      else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==VendorName){
      VendorNumber = table.getItem(1).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(49, 71);
      ReportUtils.logStep_Screenshot();  
      table.Click(49, 71);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }
      else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==VendorName){
      VendorNumber = table.getItem(2).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(49, 90);
      ReportUtils.logStep_Screenshot();
      table.Click(49, 90);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }
      else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==VendorName){
        VendorNumber = table.getItem(3).getText_2(0).OleValue.toString().trim();
      table.HoverMouse(49, 109);
      ReportUtils.logStep_Screenshot();
      table.Click(49, 109);
      ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
      }   
      
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(4000, Indicator.Text);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(4000, Indicator.Text);  
}


function AttachDocument(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//   var doc = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
   var doc = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel;
   var docCount = doc.ChildCount;
  for(var i=0;i<docCount;i++){ 
    if((doc.Child(i).isVisible())&&(doc.Child(i).Index == docCount)){
      WorkspaceUtils.waitForObj(doc.Child(i));
      ReportUtils.logStep_Screenshot("");
      doc = doc.Child(i);
      break;
    }
  }
  Sys.HighlightObject(doc);
  doc.HoverMouse();
  doc.HoverMouse();
  doc.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var attchDocument = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  Sys.HighlightObject(attchDocument);
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  ReportUtils.logStep_Screenshot();
  attchDocument.Click();
  aqUtils.Delay(4000, Indicator.Text);
  var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}

function Information(){ 
  var info = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
  info.HoverMouse();
  info.HoverMouse();
  info.HoverMouse();
  Sys.HighlightObject(info);
  info.HoverMouse();
  info.HoverMouse();
  info.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var submit = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(submit);
  submit.HoverMouse();
  submit.HoverMouse();
  submit.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}


function ApprvalInformation(){ 
        
        VendorNumber = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McValuePickerWidget.getText().OleValue.toString().trim();
        Log.Message("Vendor Number :" + VendorNumber);
        aqUtils.Delay(3000, Indicator.Text);
        ValidationUtils.verify(true,true,"Vendor Number : "+VendorNumber);
        ExcelUtils.setExcelName(workBook,"Data Management", true);
        ExcelUtils.WriteExcelSheet("Vendor Number",EnvParams.Opco,"Data Management",VendorNumber)
        ExcelUtils.WriteExcelSheet("Global Vendor Currency",EnvParams.Opco,"Data Management",Currency)

      if(ImageRepository.ImageSet0.Maximize.Exists()){
        ImageRepository.ImageSet0.Maximize.Click();
        }  
       else{ 
       var VendorApprovalpane = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;       
        Sys.HighlightObject(VendorApprovalpane);
        VendorApprovalpane.HoverMouse();
        VendorApprovalpane.Click();
          if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
          ImageRepository.ImageSet0.Maximize.Click();
        }
        
        
        var VendorApproval =  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
        //Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
        Sys.HighlightObject(VendorApproval);
        VendorApproval.HoverMouse();
        VendorApproval.Click();
          if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
           }
           var ApproverTable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//           Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
   
           var y=0;
              for(var i=0;i<ApproverTable.getItemCount();i++){   
                 var approvers="";
                  if(ApproverTable.getItem(i).getText_2(3)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
                  approvers = EnvParams.Opco+"*"+VendorNumber+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(5).OleValue.toString().trim();
                  Log.Message("Approver level :" +i+ ": " +approvers);
                  Approve_Level[y] = approvers;
                  Log.Message(Approve_Level[y])
                  y++;
                  }
              }
           TextUtils.writeLog("Finding approvers for Created Global Vendor");
        var closeCAList = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
// Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
        Sys.HighlightObject(closeCAList);
        closeCAList.HoverMouse();
        closeCAList.Click();
        ImageRepository.ImageSet0.Forward.Click();
          if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
          }
        CredentialLogin();
        var OpCo2 = ApproveInfo[0].split("*");
        ExcelUtils.setExcelName(workBook, "Server Details", true);
        var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);

        sheetName = "CreateGlobalVendor";
        if(OpCo2[2]==Project_manager){
        level = 1;
        var Approve = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
        Sys.HighlightObject(Approve)
        if(Approve.isEnabled()){ 
        Approve.HoverMouse();
        ReportUtils.logStep_Screenshot();
        Approve.Click();
        aqUtils.Delay(8000, "Waiting for Approve");;
        ValidationUtils.verify(true,true,"Vendor is Approved by "+Project_manager)
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
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
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
          if((temp.indexOf((JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()+" (")!=-1)!=-1)&&(temp1.length==2)){ 
            Client_Managt.ClickItem("|"+temp);   
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp);  
            TextUtils.writeLog("Entering into Approve Vendor by Type from To-Dos List");
            listPass = false; 
          }
      }
      if(lvl==3)
      for(var j=0;j<Client_Managt.getItemCount();j++){ 
          var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
          var temp1 = temp.split("(");
        if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
            Client_Managt.ClickItem("|"+temp);    
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp); 
            TextUtils.writeLog("Entering into Approve Vendor by Type (Substitute) from To-Dos List");
            var listPass = false;   
         }
      }  
  if(listPass){
    if(lvl==2)
          for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
              if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
                  Client_Managt.ClickItem("|"+temp);   
                  ReportUtils.logStep_Screenshot(); 
                  Client_Managt.DblClickItem("|"+temp);  
                  TextUtils.writeLog("Entering into Approve Vendor from To-Dos List");
                  listPass = false; 
                }
           }
    if(lvl==3)
        for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
          if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
              Client_Managt.ClickItem("|"+temp);    
              ReportUtils.logStep_Screenshot(); 
              Client_Managt.DblClickItem("|"+temp); 
              TextUtils.writeLog("Entering into Approve Vendor (Substitute) from To-Dos List");
              var listPass = true;   
            }
        } 
  }
}


function FinalApproveClient(VendorNum,Apvr,lvl){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

    var table = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder;
    waitForObj(table);
    Sys.HighlightObject(table);
      if(Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Visible){
      }
      else{
      var showFilter = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
      waitForObj(table);
      Sys.HighlightObject(showFilter);
      showFilter.HoverMouse();
      showFilter.HoverMouse();
      showFilter.HoverMouse();
      showFilter.Click();
      }

    var table = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//    
//      if(!linestatus) 
//    if((Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3).isVisible())
//    {   
//    var table = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//    linestatus = true;
//    }
//     if(!linestatus) 
//    if((Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite3).isVisible())
//    {   
//    var refresh = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//    linestatus = true;
//    }
//    
    var firstCell = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
    waitForObj(firstCell);
    Sys.HighlightObject(firstCell);
    firstCell.HoverMouse();
    firstCell.HoverMouse();
    firstCell.setText(VendorNum);
    aqUtils.Delay(3000, "Reading Data in table");;
    var closefilter = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    waitForObj(closefilter);
    Sys.HighlightObject(closefilter);
    closefilter.HoverMouse();
    closefilter.HoverMouse(); 
    closefilter.HoverMouse();
    closefilter.HoverMouse(); 
      var flag=false;
      for(var v=0;v<table.getItemCount();v++){ 
        if(table.getItem(v).getText_2(0).OleValue.toString().trim()==VendorNum){ 
          flag=true;    
          break;
        }
        else{ 
          table.Keys("[Down]");
        }
      }

  ValidationUtils.verify(flag,true,"Created Vendor is available in Approval List");
  TextUtils.writeLog("Created Vendor is available in Approval List");
      if(flag){ 
      closefilter.HoverMouse();
      ReportUtils.logStep_Screenshot();
      closefilter.Click();
      aqUtils.Delay(5000, Indicator.Text);

      var Approve = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      Sys.HighlightObject(Approve)
      if(Approve.isEnabled()){ 
      Approve.HoverMouse();
      ReportUtils.logStep_Screenshot();
      Approve.Click();
      aqUtils.Delay(8000, "Waiting To Approve");;
      ValidationUtils.verify(true,true,"Global Vendor is Approved by "+Apvr)
      aqUtils.Delay(8000, Indicator.Text);;
      TextUtils.writeLog("Global Vendor is Approved by "+Apvr);
      if(Approve_Level.length==lvl+1){
//      var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()).SWTObject("Label", "*");
//      Log.Message(label.getText());
//      var lab = label.getText().OleValue.toString().trim();
//      var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//      Ok.HoverMouse(); 
//      ReportUtils.logStep_Screenshot();
//      Ok.Click(); 
      aqUtils.Delay(8000, Indicator.Text); ;
       for(var j=0;j<12;j++){ 
      if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()){ 
      var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()).SWTObject("Label", "*");
      Log.Message(label.getText());
      var lab = label.getText().OleValue.toString().trim();
      var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
      Ok.HoverMouse(); 
      ReportUtils.logStep_Screenshot();
      Ok.Click(); 
      aqUtils.Delay(8000, Indicator.Text); 
      }
      }
 
//        ExcelUtils.setExcelName(workBook,"Data Management", true);
//        ExcelUtils.WriteExcelSheet("Global Vendor",EnvParams.Opco,"Data Management",VendorNum)
        TextUtils.writeLog("Global Vendor Number :"+ VendorNum); 
       var ClientApproval = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
       Sys.HighlightObject(ClientApproval);
       ClientApproval.HoverMouse();
       ClientApproval.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
       if(ImageRepository.ImageSet.Maximize.Exists()){
      ImageRepository.ImageSet.Maximize.Click();
      }
       var ClientApproval = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
       Sys.HighlightObject(ClientApproval);
       ClientApproval.HoverMouse();
       ClientApproval.Click();
       if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
       }
         var ApproverTable = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
        Sys.HighlightObject(ApproverTable);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
           }
        ReportUtils.logStep_Screenshot();
            for(var i=0;i<ApproverTable.getItemCount();i++){   
        var approvers="";
        if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
        ValidationUtils.verify(true,false,"Created Vendor is not Approved")
        }
}
        var closeApproval = Aliases.Maconomy.GVendor.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel2.TabControl;
        Sys.HighlightObject(closeApproval);
       closeApproval.HoverMouse();
       closeApproval.Click();
       if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
         }
       ImageRepository.ImageSet.Forward.Click();
       if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
       }
       var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
        menuBar.Click();
        
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Vendor Number",EnvParams.Opco,"Data Management",VendorNum)
  ExcelUtils.WriteExcelSheet("Global Vendor Currency",EnvParams.Opco,"Data Management",Currency)
      }
        ValidationUtils.verify(true,true,"Global Vendor is Approved by "+Apvr)  
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
  aqUtils.Delay(4000, Indicator.Text);;
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



function IndiaSpecific(){ 
  aqUtils.Delay(7000, Indicator.Text);
  var indiaspec = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.invoicehis;
Sys.HighlightObject(indiaspec);
var Start = StartwaitTime();
var waitTime = true;
var Difference = 0;
while(waitTime)
if(Difference<61){
if((indiaspec.isEnabled())&&(indiaspec.text=="India Specific")){
Sys.HighlightObject(indiaspec);
indiaspec.HoverMouse();
indiaspec.Click();
waitTime = false;
}
else{ 
var End = EndTime();
Difference = End - Start;
}
}
else{
ValidationUtils.verify(true,false,"Screen is not Responding more than a minute");
}
  var vendortype = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  var Statecode = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
  var GSTVendorType = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  var PermanentEstablishment = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
  var TDSapplicable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McValuePickerWidget;
  var TDSsection =  NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Composite2.McValuePickerWidget;
  var method = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Composite.McValuePickerWidget;
  
  if(Vendortype!=""){
  Sys.HighlightObject(vendortype);
  vendortype.HoverMouse();
  vendortype.Click();
  WorkspaceUtils.SearchByValue(vendortype,"Option",Vendortype,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text);  
  
  if(StateCode!=""){
  Sys.HighlightObject(Statecode);
  Statecode.HoverMouse();
  Statecode.Click();
  DropDownList(StateCode,"State Code")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  if(GSTVendor!=""){
  Sys.HighlightObject(GSTVendorType);
  GSTVendorType.HoverMouse();
  GSTVendorType.Click();
  WorkspaceUtils.SearchByValue(GSTVendorType,"Local Specification 7",GSTVendor,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text); 
  
  PermanentEstablishment.Keys("yes");
  aqUtils.Delay(2000, Indicator.Text);
  
  if(TDSAplicable!=""){
  Sys.HighlightObject(TDSapplicable);
  TDSapplicable.HoverMouse();
  TDSapplicable.Click();
  WorkspaceUtils.SearchByValue(TDSapplicable,"Option",TDSAplicable,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text);  
 
    var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
    Sys.HighlightObject(save);
    save.HoverMouse();
    save.Click();
    aqUtils.Delay(2000, Indicator.Text);
    
  if(Section!=""){
  Sys.HighlightObject(TDSsection);
  TDSsection.HoverMouse();
  TDSsection.Click();
  WorkspaceUtils.SearchByValue(TDSsection,"Local Specification 8",Section,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text);  
  
  if(Method!=""){
  Sys.HighlightObject(method);
  method.HoverMouse();
  method.Click();
  WorkspaceUtils.SearchByValue(method,"Option",Method,"Name");
  }
  aqUtils.Delay(4000, Indicator.Text);
  save.Click();
  aqUtils.Delay(2000, Indicator.Text); 
}
 