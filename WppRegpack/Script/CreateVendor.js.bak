//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
 
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateVendor";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var vendorName,strt1,PostalCode,City,country,vendorlan,taxcode,companyReg,currency,vendorgrp,controlAct,bfc,bankName,iban,swift,AccountNo,sortcode,company,attn,mail,Remail,lan,phone,fax,Paymentmode,payterm,Comtaxcode,level1Tax,estimateFirstOrder,estimateAnnualspend ="";
var VendorNumber = "";


function VendorCreation(){ 
Indicator.PushText("waiting for window to open");
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
sheetName = "CreateVendor";
Language = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
vendorName,strt1,PostalCode,City,country,vendorlan,taxcode,companyReg,currency,vendorgrp,controlAct,bfc,bankName,iban,swift,AccountNo,sortcode,company,attn,mail,Remail,lan,phone,fax,Paymentmode,payterm,Comtaxcode,level1Tax,estimateFirstOrder,estimateAnnualspend ="";
VendorNumber = "";
Approve_Level = [];

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Log.Message(EnvParams.Opco)
Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Vendor Creation started::"+STIME);
getDetails();
goToJobMenuItem(); 
vendorSearch();
globalVendor();
Global_vendor_MasterData_1();
Global_vendor_MasterData_2();
policy();
gotoCreatedVendor();
Vendors();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
Log.Message(temp[3])
todo(temp[3],i);
FinalApproveVendor(temp[0],temp[1],temp[2],i);
}
////FinalApproveClient();
}

function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
vendorName = ExcelUtils.getRowDatas("Vendor Name",EnvParams.Opco)
if((vendorName==null)||(vendorName=="")){ 
ValidationUtils.verify(false,true,"Vendor Name is Needed to Create a Vendor");
}

strt1 = ExcelUtils.getRowDatas("Street 1",EnvParams.Opco)
if((strt1==null)||(strt1=="")){ 
ValidationUtils.verify(false,true,"Street 1 is Needed to Create a Vendor");
}

PostalCode = ExcelUtils.getRowDatas("Postal Code",EnvParams.Opco)
if((PostalCode==null)||(PostalCode=="")){ 
ValidationUtils.verify(false,true,"Postal Code is Needed to Create a Vendor");
}

City = ExcelUtils.getRowDatas("City",EnvParams.Opco)
if((City==null)||(City=="")){ 
ValidationUtils.verify(false,true,"City is Needed to Create a Vendor");
}

country = ExcelUtils.getRowDatas("Country",EnvParams.Opco)
if((country==null)||(country=="")){ 
ValidationUtils.verify(false,true,"Country is Needed to Create a Vendor");
}
vendorlan = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((vendorlan==null)||(vendorlan=="")){ 
ValidationUtils.verify(false,true,"Language is Needed to Create a Vendor");
}
taxcode = ExcelUtils.getRowDatas("Tax No.",EnvParams.Opco)
if((taxcode==null)||(taxcode=="")){ 
ValidationUtils.verify(false,true,"Tax No. is Needed to Create a Vendor");
}

companyReg = ExcelUtils.getRowDatas("Company Reg. No.",EnvParams.Opco)
if((companyReg==null)||(companyReg=="")){ 
ValidationUtils.verify(false,true,"Company Reg. No. is Needed to Create a Vendor");
}
currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((currency==null)||(currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Create a Vendor");
}

vendorgrp = ExcelUtils.getRowDatas("Vendor Group",EnvParams.Opco)
if((vendorgrp==null)||(vendorgrp=="")){ 
ValidationUtils.verify(false,true,"Vendor Group is Needed to Create a Vendor");
}
controlAct = ExcelUtils.getRowDatas("Control Account",EnvParams.Opco)
if((controlAct==null)||(controlAct=="")){ 
ValidationUtils.verify(false,true,"Control Account is Needed to Create a Vendor");
}

bfc = ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)
if((bfc==null)||(bfc=="")){ 
ValidationUtils.verify(false,true,"Counter Party BFC is Needed to Create a Vendor");
}

bankName = ExcelUtils.getRowDatas("Bank Account Name",EnvParams.Opco)
if((bankName==null)||(bankName=="")){ 
ValidationUtils.verify(false,true,"Bank Account Name is Needed to Create a Vendor");
}

iban = ExcelUtils.getRowDatas("IBAN",EnvParams.Opco)
if((iban==null)||(iban=="")){ 
ValidationUtils.verify(false,true,"IBAN is Needed to Create a Vendor");
}

swift = ExcelUtils.getRowDatas("SWIFT",EnvParams.Opco)
if((swift==null)||(swift=="")){ 
ValidationUtils.verify(false,true,"SWIFT is Needed to Create a Vendor");
}

AccountNo = ExcelUtils.getRowDatas("Bank Acct. No.",EnvParams.Opco)
if((AccountNo==null)||(AccountNo=="")){ 
ValidationUtils.verify(false,true,"Bank Acct. No. is Needed to Create a Vendor");
}

sortcode = ExcelUtils.getRowDatas("Sort Code / ABA No.",EnvParams.Opco)
if((sortcode==null)||(sortcode=="")){ 
ValidationUtils.verify(false,true,"Sort Code / ABA No. is Needed to Create a Vendor");
}

company = ExcelUtils.getRowDatas("Company No.",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company No. is Needed to Create a Vendor");
}
attn = ExcelUtils.getRowDatas("Attn.",EnvParams.Opco)
if((attn==null)||(attn=="")){ 
ValidationUtils.verify(false,true,"Attn. is Needed to Create a Vendor");
}
mail = ExcelUtils.getRowDatas("E-mail",EnvParams.Opco)
if((mail==null)||(mail=="")){ 
ValidationUtils.verify(false,true,"E-mail is Needed to Create a Vendor");
}

Remail = ExcelUtils.getRowDatas("Remittance Email",EnvParams.Opco)
if((Remail==null)||(Remail=="")){ 
ValidationUtils.verify(false,true,"Remittance Email is Needed to Create a Vendor");
}

lan = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((lan==null)||(lan=="")){ 
ValidationUtils.verify(false,true,"Language is Needed to Create a Vendor");
}

phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
if((phone==null)||(phone=="")){ 
ValidationUtils.verify(false,true,"Phone is Needed to Create a Vendor");
}

fax = ExcelUtils.getRowDatas("Fax",EnvParams.Opco)
if((fax==null)||(fax=="")){ 
ValidationUtils.verify(false,true,"Fax is Needed to Create a Vendor");
}

Paymentmode = ExcelUtils.getRowDatas("Payment Mode",EnvParams.Opco)
if((Paymentmode==null)||(Paymentmode=="")){ 
ValidationUtils.verify(false,true,"Payment Mode is Needed to Create a Vendor");
}

payterm = ExcelUtils.getRowDatas("Payment Terms",EnvParams.Opco)
if((payterm==null)||(payterm=="")){ 
ValidationUtils.verify(false,true,"Payment Terms is Needed to Create a Vendor");
}

Comtaxcode = ExcelUtils.getRowDatas("Company Tax Code",EnvParams.Opco)
if((Comtaxcode==null)||(Comtaxcode=="")){ 
ValidationUtils.verify(false,true,"Company Tax Code is Needed to Create a Vendor");
}

level1Tax = ExcelUtils.getRowDatas("Level 1 Tax Derivation",EnvParams.Opco)
if((level1Tax==null)||(level1Tax=="")){ 
ValidationUtils.verify(false,true,"Level 1 Tax Derivation is Needed to Create a Vendor");
}

estimateFirstOrder = ExcelUtils.getRowDatas("What is the estimated value of first order with this supplier in the supplier’s currency?",EnvParams.Opco)
if((estimateFirstOrder==null)||(estimateFirstOrder=="")){ 
ValidationUtils.verify(false,true,"What is the estimated value of first order with this supplier in the supplier’s currency? is Needed to Create a Vendor");
}
estimateAnnualspend = ExcelUtils.getRowDatas("What is the estimated annual spend with this supplier in the supplier’s currency?",EnvParams.Opco)
if((estimateAnnualspend==null)||(estimateAnnualspend=="")){ 
ValidationUtils.verify(false,true,"What is the estimated annual spend with this supplier in the supplier’s currency? is Needed to Create a Vendor");
}

}

function goToJobMenuItem(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountPayable.Exists()){
ImageRepository.ImageSet.AccountPayable.Click();// GL
}
else if(ImageRepository.ImageSet.AccountPayable2.Exists()){
ImageRepository.ImageSet.AccountPayable2.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
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

function vendorSearch_Address(){ 
  aqUtils.Delay(4000, Indicator.Text);;
  var Company_No = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget.getText().OleValue.toString().trim()
  if(Company_No!="Company No.")
  ValidationUtils.verify(false,true,"Company No field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Company No field is available in Maconomy for Vendor Creation");
  
  var Currency = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget.getText().OleValue.toString().trim()
  if(Currency!="Currency")
  ValidationUtils.verify(false,true,"Currency field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Currency field is available in Maconomy for Vendor Creation");
  
  var Vendor_name = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget2.getText().OleValue.toString().trim()
  if(Vendor_name!="Vendor Name")
  ValidationUtils.verify(false,true,"Vendor Name field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Vendor Name field is available in Maconomy for Vendor Creation");
  
}

function vendorSearch(){ 
vendorSearch_Address();

  var Company_No = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  if(company!=""){
  Company_No.Click();
  WorkspaceUtils.SearchByValue(Company_No,"Company",company,"Company Number");
    }
    
  var Currency = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
  if(currency!=""){
  Currency.Keys(" ");
  aqUtils.Delay(4000, Indicator.Text);
  Currency.Click();
  WorkspaceUtils.DropDownList(currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  var Vendor_name = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
  if(vendorName!=""){
  Vendor_name.setText(vendorName+" "+STIME);
  ValidationUtils.verify(true,true,"Client Name Entered in Global Client Data 1/2");
     }
     
  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.RemarksSave;
  save.HoverMouse();
  ReportUtils.logStep_Screenshot();
  save.Click();
  aqUtils.Delay(3000, Indicator.Text);
}


function globalVendor(){ 
  var gVendor = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  gVendor.Click();
  aqUtils.Delay(4000, Indicator.Text);
  var newVendor = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  newVendor.HoverMouse();
  ReportUtils.logStep_Screenshot();  
  newVendor.Click();
  aqUtils.Delay(4000, Indicator.Text);
}


function Global_vendor_MasterData_1_Address(){ 
  aqUtils.Delay(4000, Indicator.Text);;
  var Country = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget.getText().OleValue.toString().trim()
  if(Country!="Country")
  ValidationUtils.verify(false,true,"Country field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Country field is available in Maconomy for Vendor Creation");
  
  var Tax_No = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget2.getText().OleValue.toString().trim()
  if(Tax_No!="Tax No.")
  ValidationUtils.verify(false,true,"Tax No field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Tax No field is available in Maconomy for Vendor Creation");
  
  var Compy_Reg_no = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget2.getText().OleValue.toString().trim()
  if(Compy_Reg_no!="Comp. Reg. No.")
  ValidationUtils.verify(false,true,"Comp. Reg. No. field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Comp. Reg. No. field is available in Maconomy for Vendor Creation");
  
  var Currency = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McTextWidget2.getText().OleValue.toString().trim()
  if(Currency!="Currency")
  ValidationUtils.verify(false,true,"Currency field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Currency field is available in Maconomy for Vendor Creation");
  
  var vendor_grp = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McTextWidget.getText().OleValue.toString().trim()
  if(vendor_grp!="Vendor Group")
  ValidationUtils.verify(false,true,"Vendor Group field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Vendor Group field is available in Maconomy for Vendor Creation");
  
  var control_Acc = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McTextWidget.getText().OleValue.toString().trim()
  if(control_Acc!="Control Account")
  ValidationUtils.verify(false,true,"Control Account field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Control Account field is available in Maconomy for Vendor Creation");
  
  var party_BFC = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.McTextWidget.getText().OleValue.toString().trim()
  if(party_BFC!="Counter Party BFC")
  ValidationUtils.verify(false,true,"Counter Party BFC field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Counter Party BFC field is available in Maconomy for Vendor Creation");
  
  var BankAccountName = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.McTextWidget2.getText().OleValue.toString().trim()
  if(BankAccountName!="Bank Account Name")
  ValidationUtils.verify(false,true,"Bank Account Name field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Bank Account Name field is available in Maconomy for Vendor Creation");
  
  var IBAN1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.McTextWidget2.getText().OleValue.toString().trim()
  if(IBAN1!="IBAN")
  ValidationUtils.verify(false,true,"IBAN field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"IBAN field is available in Maconomy for Vendor Creation");
  
  var swift1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.McTextWidget2.getText().OleValue.toString().trim()
  if(swift1!="SWIFT")
  ValidationUtils.verify(false,true,"SWIFT field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"SWIFT field is available in Maconomy for Vendor Creation");
  
  var AccountNumber = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.McTextWidget2.getText().OleValue.toString().trim()
  if(AccountNumber!="Bank Acct. No.")
  ValidationUtils.verify(false,true,"Bank Acct. No. field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Bank Acct. No. field is available in Maconomy for Vendor Creation");
  
  var sortcode1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite15.McTextWidget2.getText().OleValue.toString().trim()
  if(sortcode1!="Sort Code / ABA No.")
  ValidationUtils.verify(false,true,"Sort Code / ABA No. field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Sort Code / ABA No. field is available in Maconomy for Vendor Creation");
  
}

function Global_vendor_MasterData_1(){ 
  Global_vendor_MasterData_1_Address();
  ReportUtils.logStep("INFO","Entering data in Global Vendor Master Data 1/1");
  var Vendor_name = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
    
  if(vendorName!=""){
  vendorName = vendorName+" "+STIME;
  Vendor_name.setText(vendorName);
  ValidationUtils.verify(true,true,"Vendor Name Entered in Global Vendor Master Data 1/2");
     }
     
  var street1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  if(strt1!=""){
  street1.setText(strt1);
  ValidationUtils.verify(true,true,"Street1 Entered in Global Vendor Master Data 1/2");
     }
     
  var Postal_code = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.McValuePickerWidget;
  if(PostalCode!=""){
  Postal_code.setText(PostalCode);
  ValidationUtils.verify(true,true,"Postal Code Entered in Global Vendor Master Data 1/2");
   }
  
  aqUtils.Delay(2000, Indicator.Text);;
  var PostalCity = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.McValuePickerWidget2;
  if(City!=""){
  PostalCity.setText(City);
  ValidationUtils.verify(true,true,"City Entered in Global Vendor Master Data 1/2");
   }
   
  var Country = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
  if(country!=""){
  Country.Click();
  WorkspaceUtils.DropDownList(country,"Country")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  var Tax_No = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
  if(taxcode!=""){
  Tax_No.setText(taxcode);
  ValidationUtils.verify(true,true,"Tax No Entered in Global Vendor Master Data 1/2");
     }
     
  var Compy_Reg_no = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget;
  if(companyReg!=""){
  Compy_Reg_no.setText(companyReg);
  ValidationUtils.verify(true,true,"Company Registration No Entered in Global Vendor Master Data 1/2");
     }
     
  var Currency = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McPopupPickerWidget;
  if(currency!=""){
  Currency.Click();
  WorkspaceUtils.DropDownList(currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);  
  var vendor_grp = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McPopupPickerWidget;
  if(vendorgrp!=""){
  vendor_grp.Click();
  WorkspaceUtils.DropDownList(vendorgrp,"Vendor Group")
  }
  aqUtils.Delay(2000, Indicator.Text);
  var control_Acc = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McPopupPickerWidget;
  if(controlAct!=""){
  control_Acc.Click();
  WorkspaceUtils.DropDownList(controlAct,"Control Account")
  }
  aqUtils.Delay(2000, Indicator.Text);

  var party_BFC = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.McValuePickerWidget;
  if(bfc!=""){
  party_BFC.Click();
  WorkspaceUtils.SearchByValue(party_BFC,"Counter Party BFC",bfc,"Counter Party BFC");
    }
    
  var BankAccountName = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.McTextWidget;
  if(bankName!=""){
  BankAccountName.setText(bankName);
  ValidationUtils.verify(true,true,"Bank Account Name Entered in Global Vendor Master Data 1/2");
     }
     
  var IBAN1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.McTextWidget;
  if(iban!=""){
  IBAN1.setText(iban);
  ValidationUtils.verify(true,true,"IBAN Entered in Global Vendor Master Data 1/2");
     }
     
  var swift1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.McTextWidget;
  if(swift!=""){
  swift1.setText(swift);
  ValidationUtils.verify(true,true,"SWIFT Entered in Global Vendor Master Data 1/2");
     }
     
  var AccountNumber = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.McTextWidget;
  if(AccountNo!=""){
  AccountNumber.setText(AccountNo);
  ValidationUtils.verify(true,true,"Account No Entered in Global Vendor Master Data 1/2");
     }
     
  var sortcode1 = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite15.McTextWidget;
  if(sortcode!=""){
  sortcode1.setText(sortcode);
  ValidationUtils.verify(true,true,"Sort Code Entered in Global Vendor Master Data 1/2");
     }
  
  var next = Aliases.Maconomy.Shell7.Composite.Composite.Composite2.Composite.Composite.Button;
  next.HoverMouse();
  ReportUtils.logStep_Screenshot();
  next.Click();
}


function Global_vendor_MasterData_2_Address(){ 
  aqUtils.Delay(4000, Indicator.Text);;
  var Company_No = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget2.getText().OleValue.toString().trim()
  if(Company_No!="Company No.")
  ValidationUtils.verify(false,true,"Company No. field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Company No. field is available in Maconomy for Vendor Creation");
  
  var Attn = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget2.getText().OleValue.toString().trim()
  if(Attn!="Attn.")
  ValidationUtils.verify(false,true,"Attn. field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Attn. field is available in Maconomy for Vendor Creation");
  
  var Email = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite16.McTextWidget2.getText().OleValue.toString().trim()
  if(Email!="E-mail")
  ValidationUtils.verify(false,true,"E-mail field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"E-mail field is available in Maconomy for Vendor Creation");
  
  var ReEmail = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite17.McTextWidget2.getText().OleValue.toString().trim()
  if(ReEmail!="Remittance Email")
  ValidationUtils.verify(false,true,"Remittance Email field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Remittance Email field is available in Maconomy for Vendor Creation");
  
  var language = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite18.McTextWidget.getText().OleValue.toString().trim()
  if(language!="Language")
  ValidationUtils.verify(false,true,"Language field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Language field is available in Maconomy for Vendor Creation");
  
  var Phone = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget2.getText().OleValue.toString().trim()
  if(Phone!="Phone")
  ValidationUtils.verify(false,true,"Phone field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Phone field is available in Maconomy for Vendor Creation");
  
  var Company_Tax_Code = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget2.getText().OleValue.toString().trim()
  if(Company_Tax_Code!="Company Tax Code")
  ValidationUtils.verify(false,true,"Company Tax Code field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Company Tax Code field is available in Maconomy for Vendor Creation");
  
  var Level_1_Tax_Derivation = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget2.getText().OleValue.toString().trim()
  if(Level_1_Tax_Derivation!="Level 1 Tax Derivation")
  ValidationUtils.verify(false,true,"Level 1 Tax Derivation field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Level 1 Tax Derivation field is available in Maconomy for Vendor Creation");
  
  var Payment_Terms = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McTextWidget2.getText().OleValue.toString().trim()
  if(Payment_Terms!="Payment Terms")
  ValidationUtils.verify(false,true,"Payment Terms field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Payment Terms field is available in Maconomy for Vendor Creation");
  
  var Vendor_Payment_Mode = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McTextWidget.getText().OleValue.toString().trim()
  if(Vendor_Payment_Mode!="Payment Mode")
  ValidationUtils.verify(false,true,"Payment Mode field is missing in Maconomy for Vendor Creation");
  else
  ValidationUtils.verify(true,true,"Payment Mode field is available in Maconomy for Vendor Creation");
  
  
}

function Global_vendor_MasterData_2(){ 

  Global_vendor_MasterData_2_Address();
  var Company_No = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(3000, Indicator.Text);  
  
  if(company!=""){
  Company_No.Click();
  WorkspaceUtils.SearchByValue(Company_No,"Company",company,"Company Number");
    }
    
  var Attn = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  if(attn!=""){
  Attn.setText(attn);
  ValidationUtils.verify(true,true,"Attn Entered in Global Vendor Master Data 2/2");
     }
     
  var Email = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite16.McTextWidget;
  if(mail!=""){
  Email.setText(mail);
  ValidationUtils.verify(true,true,"Email Entered in Global Vendor Master Data 2/2");
     }
     
  var ReEmail = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite17.McTextWidget;
  if(Remail!=""){
  ReEmail.setText(Remail);
  ValidationUtils.verify(true,true,"Remittance Email Entered in Global Vendor Master Data 2/2");
     }
     
  var language = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite18.McPopupPickerWidget;
  if(lan!=""){
  language.Click();
  WorkspaceUtils.DropDownList(lan,"Language")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
 var Phone = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget; 
 if(phone!=""){
 Phone.setText(phone);
 ValidationUtils.verify(true,true,"Phone Number is Entered in Global Vendor Master Data 2/2");
     }
     
 var Company_Tax_Code = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McPopupPickerWidget; 
 if(Comtaxcode!=""){
  Company_Tax_Code.Click();  
  aqUtils.Delay(5000, Indicator.Text);;
  WorkspaceUtils.DropDownList(Comtaxcode,"Company Tax Code"); 
     }


  var Level_1_Tax_Derivation = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McValuePickerWidget; 
  Level_1_Tax_Derivation.setText("-");
  
  var Payment_Terms = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McPopupPickerWidget; 
  if(payterm !=""){
  Payment_Terms.Click();  
  aqUtils.Delay(5000, Indicator.Text);;
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(payterm,"Payment Terms"); 
     }
     
  var Vendor_Payment_Mode = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.McValuePickerWidget; 
  if(Paymentmode!=""){
  Vendor_Payment_Mode.Click();
  WorkspaceUtils.SearchByValue(Vendor_Payment_Mode,"Payment Mode",Paymentmode,"Payment Mode");
         }
         
  var next = Aliases.Maconomy.Shell7.Composite.Composite.Composite2.Composite.Composite.Button;
  next.HoverMouse();
  ReportUtils.logStep_Screenshot();
  next.Click();
     
}


function policy(){ 
  aqUtils.Delay(3000, Indicator.Text);
  var confirm = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite19.McPopupPickerWidget;
  confirm.Keys("Yes");
  aqUtils.Delay(3000, Indicator.Text);
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(2000, Indicator.Text);
  
  var next = Aliases.Maconomy.Shell7.Composite.Composite.Composite2.Composite.Composite.Button;
  next.HoverMouse();
  ReportUtils.logStep_Screenshot();
  next.Click();
  
  
  var Document = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(4000, Indicator.Text);
  Document.setText("Yes");
  
  var WPPPreferredSupplier = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite16.McPopupPickerWidget;
  WPPPreferredSupplier.Keys("Yes");

  aqUtils.Delay(2000, Indicator.Text);
  var  newSupplierCreated = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite17.McTextWidget
  newSupplierCreated.setText("Yes");
  
  var WPPAdvisorpaymentpolicy = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
  WPPAdvisorpaymentpolicy.Keys("Yes");

  aqUtils.Delay(2000, Indicator.Text);
  var supplierCapacity = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McPopupPickerWidget;
  supplierCapacity.Keys("Yes");

  aqUtils.Delay(2000, Indicator.Text);
  var Document = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McTextWidget;
  Document.setText("Yes");
  
  var personalRelationships = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.McPopupPickerWidget;
  personalRelationships.Keys("Yes");

  aqUtils.Delay(2000, Indicator.Text);
  var  identifiedRelationship = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.McTextWidget;
  identifiedRelationship.setText("Yes");
  var  firstorder = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.McTextWidget;
  firstorder.setText(estimateFirstOrder);
  var  annualspend = Aliases.Maconomy.Shell7.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.McTextWidget;
  annualspend.setText(estimateAnnualspend);
  
  var create = Aliases.Maconomy.Shell7.Composite.Composite.Composite2.Composite.Button;
  create.HoverMouse();
  ReportUtils.logStep_Screenshot();
  create.Click();
  aqUtils.Delay(5000, Indicator.Text);
  var label = Aliases.Maconomy.Shell8.Label;
  var lab = label.getText().OleValue.toString().trim();
  ReportUtils.logStep("INFO",lab)
  Log.Message(lab);
  var OK = Aliases.Maconomy.Shell8.Composite.Button;
  OK.HoverMouse();
  ReportUtils.logStep_Screenshot();
  OK.Click();
  aqUtils.Delay(2000, Indicator.Text);
}



function gotoCreatedVendor()
{
  var blocked = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  blocked.Click();
  aqUtils.Delay(5000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var vendor_Name = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  vendor_Name.Click();
  vendor_Name.Keys(vendorName);
  aqUtils.Delay(5000, Indicator.Text);
  
  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==vendorName){
  VendorNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==vendorName){
  VendorNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==vendorName){
  VendorNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==vendorName){
  VendorNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  }
  
  aqUtils.Delay(5000, Indicator.Text);
}

function Vendors(){ 
  aqUtils.Delay(8000, Indicator.Text);;
  var document = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  document.Click();
  aqUtils.Delay(4000, Indicator.Text);
  var attchDocument = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
  attchDocument.HoverMouse();
  ReportUtils.logStep_Screenshot();
  attchDocument.Click();
  aqUtils.Delay(4000, Indicator.Text);;
  var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  information.Click();
  aqUtils.Delay(5000, Indicator.Text);
  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.RemarksSave;
  save.HoverMouse();
  ReportUtils.logStep_Screenshot(); 
  save.Click();
  aqUtils.Delay(3000, Indicator.Text);
  var submit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.SingleToolItemControl;
  submit.HoverMouse();
  ReportUtils.logStep_Screenshot();
  submit.Click();
  aqUtils.Delay(8000, Indicator.Text);
  
  var vendorapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.Approvals;
  vendorapproval.Click();
  aqUtils.Delay(2000, Indicator.Text);;
  if(ImageRepository.ImageSet.Maximize.Exists()){
  ImageRepository.ImageSet.Maximize.Click();
  }

  aqUtils.Delay(3000, Indicator.Text);;
  var AllApproved = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl2;
  AllApproved.HoverMouse();
  ReportUtils.logStep_Screenshot();
  AllApproved.Click();
  aqUtils.Delay(4000, Indicator.Text);;
  var y =0 ;
  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
  for(var i=0;i<ApproverTable.getItemCount();i++){   
  var approvers="";
  if(ApproverTable.getItem(i).getText_2(5).OleValue.toString().trim()!="Approved"){
  approvers = EnvParams.Opco+"*"+VendorNumber+"*"+ApproverTable.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
  Log.Message("Approver level :" +i+ ": " +approvers);
  Approve_Level[y] = approvers;
  y++;
  }
  }
  var moreinfo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel2.TabControl;
  moreinfo.Click();
  aqUtils.Delay(3000, Indicator.Text);;
  if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
  }
  aqUtils.Delay(4000, Indicator.Text);;
  var OpCo1 = EnvParams.Opco;
  ExcelUtils.setExcelName(workBook, "Server Details", true);
  var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
  var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
  if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
  level = 1;
  var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.SingleToolItemControl;
  Sys.HighlightObject(Approve)
  if(Approve.isEnabled()){ 
  Approve.HoverMouse();
  ReportUtils.logStep_Screenshot();
  Approve.Click();  
  aqUtils.Delay(8000, Indicator.Text);; 
  }
  }

}

function todo(lvl,clientLvl){ 
  var toDo = Aliases.Maconomy.Shell.Composite.Composite.Composite.TodoGrid.PTabFolder.TabFolderPanel.ToDo;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
if(clientLvl==(ApproveInfo.length-1)){
if(lvl==3){
Client_Managt.ClickItem("|Approve Vendor by Type (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Vendor by Type (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Vendor by Type (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Vendor by Type (*)");
}
}
else{ 
 if(lvl==3){
Client_Managt.ClickItem("|Approve Vendor (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Vendor (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Vendor (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Vendor (*)");
} 
}

}

function CredentialLogin(){ 
for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 

     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }
  else{ 
   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
    if(UserN){ 
      goToHR();
      UserN = false;
    }
    temp = searchNumber(Eno);
  }
//  Log.Message(temp)
  if(temp.length!=0){
    temp = temp+"*"+j;
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
  break;
  }
  }
  if((temp=="")||(temp==null))
  Log.Error("User Name is Not available for level :"+i);
//  Log.Message("Logins :"+temp);
}
WorkspaceUtils.closeAllWorkspaces();

}


function FinalApproveVendor(comID,cltID,apvr,clientLvl){ 
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid;
var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.McValuePickerWidget;
firstCell.setText(cltID);
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter;
  
aqUtils.Delay(6000, Indicator.Text);;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==cltID){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
 
ValidationUtils.verify(flag,true,"Created Client is available in system");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();

aqUtils.Delay(2000, Indicator.Text);;
if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}else{ 
var sideBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
sideBar.Click();
aqUtils.Delay(2000, Indicator.Text);;
ImageRepository.ImageSet.Maximize.Click();
}

var AllApproved = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
AllApproved.Click();
aqUtils.Delay(8000, Indicator.Text); ;
ReportUtils.logStep_Screenshot();
if(clientLvl==(ApproveInfo.length-1)){
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Vendor Number",EnvParams.Opco,"Data Management",cltID)
}
ValidationUtils.verify(true,true,"Vendor is approved by :"+apvr);
var closeInfor = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.listPurchaseOrder.TabControl;
closeInfor.Click();
aqUtils.Delay(2000, Indicator.Text);;
}
  }


}
