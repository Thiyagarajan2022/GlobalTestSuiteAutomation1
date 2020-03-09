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
var sheetName = "AmendCompanyVendor";
var VendorNo,email,PhoneNum,PaymentMode,CompanyReg,Currency ="";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;

function AmendCompanyVendor(){ 
  TextUtils.writeLog("Amend Gloabl Vendor Started");
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("Central Team - Vendor Account Management","Username")
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
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "AmendCompanyVendor";
VendorNo,email,VendorName ="";
ExcelUtils.setExcelName(workBook, sheetName, true);

Currency = ExcelUtils.getRowDatas("currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Amend Company Vendor");
}
email = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
Log.Message(email)
if((email==null)||(email=="")){ 
ValidationUtils.verify(false,true,"Email is Needed to Amend Company Vendor");
}
PhoneNum = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
if((PhoneNum==null)||(PhoneNum=="")){ 
ValidationUtils.verify(false,true,"Phone Number is Needed to Amend Company Vendor");
}
PaymentMode = ExcelUtils.getRowDatas("Payment Mode",EnvParams.Opco)
if((PaymentMode==null)||(PaymentMode=="")){ 
ValidationUtils.verify(false,true,"Payment Mode is Needed to Amend Company Vendor");
}
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
  if((VendorNo=="")||(VendorNo==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
  }
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Amend Company Vendor");
}



Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Amend Block Vendor started::"+STIME);
gotoMenu();
gotoVendorSearch();
companyVendor();
goToVendor();
WorkspaceUtils.closeAllWorkspaces();
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
 var CompanyNumber = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget; 
  waitForObj(CompanyNumber);
  Sys.HighlightObject(CompanyNumber)
  CompanyNumber.Click();
  WorkspaceUtils.SearchByValue(CompanyNumber,"Company",EnvParams.Opco,"Company Number");

 var curr = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
//  aqUtils.Delay(2000, Indicator.Text);
  
 var VendorNumber = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(VendorNo!=""){
  VendorNumber.Click();
  WorkspaceUtils.VPWSearchByValue(VendorNumber,"Vendor",VendorNo,"Vendor Number");
    }
    
 var Vendor_Name = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  Vendor_Name.HoverMouse();
 Sys.HighlightObject(Vendor_Name);  
  Vendor_Name.setText("*");
 
 var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
 save.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  TextUtils.writeLog("Company Number, Vendor Number, Currency has entered and Saved in Vendor Search screen");
}

function companyVendor(){    
  var active = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  waitForObj(active);
  active.HoverMouse();
  Sys.HighlightObject(active);
   active.Click();
   aqUtils.Delay(3000, "Reading from Company Vendor table");
  var table = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(51, 60);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 60);
  ValidationUtils.verify(true,true,"Amend Company Vendor is available in maconomy");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(51, 79);
  ReportUtils.logStep_Screenshot();  
  table.Click(51, 79);
  ValidationUtils.verify(true,true,"Amend Company Vendor is available in maconomy");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(51, 98);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 98);
  ValidationUtils.verify(true,true,"Amend Company Vendor is available in maconomy");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(51, 117);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 117);
  ValidationUtils.verify(true,true,"Amend Company Vendor is available in maconomy");
  }  
  aqUtils.Delay(5000, "Playback");
  TextUtils.writeLog("Amend Company Vendor is available in maconomy");
}

function goToVendor(){  

  var home = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;
   waitForObj(home);
  Sys.HighlightObject(home);
  home.HoverMouse();   
  home.Click();

  var information = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.clientpayment;
  waitForObj(information);
  Sys.HighlightObject(information);
  information.HoverMouse(); 
  information.Click();
  
 var phone = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
 var Email = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget;
   var PayMode = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
    var RemittanceEmail = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McTextWidget;
    waitForObj(phone);
  Sys.HighlightObject(phone);
  waitForObj(Email);
  Sys.HighlightObject(Email);

    var changes = false;
 
    if(PhoneNum!=""){
    if(phone.getText()!=PhoneNum){
    phone.setText(PhoneNum);
    ValidationUtils.verify(true,true,"Phone Number is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Phone Number in datasheet is as same as Value in Maconomy")
    }        
    if(PaymentMode!=""){
    if(PayMode.getText()!=PaymentMode){
      if(PaymentMode!=""){
        PayMode.Click();
        WorkspaceUtils.SearchByValue(PayMode,"Payment Mode",PaymentMode,"Name");
    }
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given CompanyReg Number in datasheet is as same as Value in Maconomy")
    } 
       
    if(email!=""){
    if(Email.getText()!=email){
    Email.setText(email);
    ValidationUtils.verify(true,true,"Email is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Email in datasheet is as same as Value in Maconomy")
    }
    
     if(email!=""){
    if(RemittanceEmail.getText()!=email){
    RemittanceEmail.setText(email);
    ValidationUtils.verify(true,true,"RemittanceEmail is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Remittance Email in datasheet is as same as Value in Maconomy")
    }
    
    if(changes){ 
  var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
  save.Click();
  aqUtils.Delay(2000, Indicator.Text);
  ReportUtils.logStep_Screenshot();
  ValidationUtils.verify(true,true,"Amend Company Vendor field are updated and saved in macanomy");
    }
    else{ 
      ValidationUtils.verify(false,true,"There is no changes happen in Maconomy screen")
    }
}


function DropDownList(value){ 
var checkMark = false;
Sys.Process("Maconomy").Refresh();
  var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  var Add_Visible4 = true;
  while(Add_Visible4){
  if(list.isEnabled()){
  Add_Visible4 = false;
      for(var i=0;i<list.getItemCount();i++){ 
        if(list.getItem(i).getText_2(0)!=null){ 
          if(list.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
            list.Keys("[Enter]");
            aqUtils.Delay(5000, Indicator.Text);;
            checkMark = true;
            ValidationUtils.verify(true,true,"Yes is selected in Blocked Status");
            break;
          }else{
            list.Keys("[Down]");
          }
          
        }else{ 
        Log.Message("i :"+i);
        Log.Message(list.getItem(i).getText_2(0).OleValue.toString().trim());
          list.Keys("[Down]");
        }
      }
  }
  }
  return checkMark;
}

 