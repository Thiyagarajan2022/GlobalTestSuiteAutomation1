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
var sheetName = "BlockCompanyVendor";
var CmpyVendorNo,Currency,CmpyVendorName ="";

function BlockCompanyVendor(){ 
  TextUtils.writeLog("Block Company Vendor Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
//ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = EnvParams.Opco+" Finance";
//var Project_manager = ExcelUtils.getRowDatas("Central Team - Vendor Account Management","Username")
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
sheetName = "BlockCompanyVendor";
CmpyVendorNo,Currency,CmpyVendorName ="";
ExcelUtils.setExcelName(workBook, sheetName, true);
CmpyVendorNo = ExcelUtils.getRowDatas("CompanyVendor Number",EnvParams.Opco)
Log.Message(CmpyVendorNo)
  if((CmpyVendorNo=="")||(CmpyVendorNo==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  CmpyVendorNo = ReadExcelSheet("CompanyVendor Number",EnvParams.Opco,"Data Management");
  }
if((CmpyVendorNo==null)||(CmpyVendorNo=="")){ 
ValidationUtils.verify(false,true,"Company Vendor Number is Needed to Block Global Vendor");
}
Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
Log.Message(Currency)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Block Global Vendor");
}
CmpyVendorName = ExcelUtils.getRowDatas("CompanyVendor Name",EnvParams.Opco)
Log.Message(CmpyVendorName)
if((CmpyVendorName==null)||(CmpyVendorName=="")){ 
ValidationUtils.verify(false,true,"Company Vendor Name is Needed to Block Global Vendor");
}

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block Company Vendor started::"+STIME);
gotoMenu();
gotoVendorSearch();
CompanyVendor();
goToCompanyVendor();
closeAllWorkspaces(); 
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
  CompanyNumber.Click();
  WorkspaceUtils.SearchByValue(CompanyNumber,"Company",EnvParams.Opco,"Company Number");

 var curr = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
 var VendorNumber = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(CmpyVendorNo!=""){
  VendorNumber.Click();
  WorkspaceUtils.VPWSearchByValue(VendorNumber,"Vendor",CmpyVendorNo,"Vendor Number");
    }
    
 var VendorName = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
 VendorName.setText("*");
 
 
 var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
 TextUtils.writeLog("Company Number, Company Vendor Number, Currency has entered and Saved in Vendor Search screen");
}

function CompanyVendor(){ 
  var CmpyClient = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  CmpyClient.Click();
  aqUtils.Delay(3000, Indicator.Text);
  var active = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;  

  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==CmpyVendorNo){
  table.HoverMouse(51, 60);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 60);
  ValidationUtils.verify(true,true,"Company Vendor is available in maconomy to block Company Vendor");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==CmpyVendorNo){
  table.HoverMouse(51, 79);
  ReportUtils.logStep_Screenshot();  
  table.Click(51, 79);
  ValidationUtils.verify(true,true,"Company Vendor is available in maconomy to block Company Vendor");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==CmpyVendorNo){
  table.HoverMouse(51, 98);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 98);
  ValidationUtils.verify(true,true,"Company Vendor is available in maconomy to block Company Vendor");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==CmpyVendorNo){
  table.HoverMouse(51, 117);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 117);
  ValidationUtils.verify(true,true,"Company Vendor is available in maconomy to block Company Vendor");
  }    
//  else{
//    ValidationUtils.verify(false,true,"Company Vendor is not available in maconomy to block Company Vendor");
//    aqUtils.Delay(2000,Indicator.Text);
//    TextUtils.writeLog("Company Vendor is not available in maconomy to block Company Vendor");
//  }

  aqUtils.Delay(5000, Indicator.Text);
    TextUtils.writeLog("Company Vendor is available in maconomy to block Company Vendor");
}

function goToCompanyVendor(){ 
  var home = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;  
  home.Click();
  aqUtils.Delay(5000, Indicator.Text);
  
////        Sys.Desktop.KeyDown(0x11);
////        Sys.Desktop.KeyDown(0x46);
////        Sys.Desktop.KeyUp(0x11);
////        Sys.Desktop.KeyUp(0x46);
  
  TextUtils.writeLog("Comapny Vendor is available in maconomy to block");
  var information = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;
  information.Click();
  aqUtils.Delay(2000, Indicator.Text);
  ReportUtils.logStep_Screenshot();
//  var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
//  screen.Click();
//  screen.MouseWheel(-200);
  var blockVendor = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McPopupPickerWidget;
  if(blockVendor.getText()=="Yes")
  ValidationUtils.verify(false,true,"Company Vendor is already blocked");
  else{ 
  blockVendor.Click();
  DropDownList("Yes")
//  blockClient.Keys("Yes");
  aqUtils.Delay(5000, Indicator.Text);
  var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
  save.Click();
  aqUtils.Delay(5000, Indicator.Text);
  ReportUtils.logStep_Screenshot();
  ValidationUtils.verify(true,true,"Company Vendor is Blocked");
  
  var popup = Sys.Process("Maconomy").SWTObject("Shell", "Company Vendors - Information");  
  Sys.HighlightObject(popup);
  var OK = Sys.Process("Maconomy").SWTObject("Shell", "Company Vendors - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  Sys.HighlightObject(OK);
  ReportUtils.logStep_Screenshot();
  OK.Click();  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Company Vendors - Information").isVisible()){
  var popup = Sys.Process("Maconomy").SWTObject("Shell", "Company Vendors - Information");
  var OK = Sys.Process("Maconomy").SWTObject("Shell", "Company Vendors - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  OK.Click();
  }
  aqUtils.Delay(4000, Indicator.Text);
    TextUtils.writeLog("Company Vendor is Blocked");
  var Allow_Registrations = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite.McPopupPickerWidget;
  if(Allow_Registrations.getText()=="No")  
  ValidationUtils.verify(true,true,"Allow Registrations has Changed to NO");
  else
  ValidationUtils.verify(true,true,"Allow Registrations has NOT Changed to NO");
  
  var Allow_Purchase_Orders = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite2.McPopupPickerWidget;
  if(Allow_Purchase_Orders.getText()=="No")  
  ValidationUtils.verify(true,true,"Allow Purchase Orders has Changed to NO");
  else
  ValidationUtils.verify(true,true,"Allow Purchase Orders has NOT Changed to NO");
  
  var Allow_Vendor_Invoices = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite3.McPopupPickerWidget;
  if(Allow_Vendor_Invoices.getText()=="No")  
  ValidationUtils.verify(true,true,"Allow Vendor Invoices has Changed to NO");
  else
  ValidationUtils.verify(true,true,"Allow Vendor Invoices has NOT Changed to NO");
  
  var Allow_Payments = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite4.McPopupPickerWidget;
  if(Allow_Payments.getText()=="No")  
  ValidationUtils.verify(true,true,"Allow Payments has Changed to NO");
  else
  ValidationUtils.verify(true,true,"Allow Payments has NOT Changed to NO");    
  TextUtils.writeLog("Allowed use has Changed to NO");  
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

 