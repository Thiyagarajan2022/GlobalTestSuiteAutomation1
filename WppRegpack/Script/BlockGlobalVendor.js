﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "BlockGlobalVendor";
var VendorNo,Currency,VendorName ="";

function BlockGlobalVendor(){ 
  TextUtils.writeLog("Block Gloabl Vendor Started");
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("Central Team - Vendor Account Management","Username")
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
sheetName = "BlockGlobalVendor";
VendorNo,Currency,VendorName ="";
ExcelUtils.setExcelName(workBook, sheetName, true);

Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Block Global Vendor");
}

VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
  if((VendorNo=="")||(VendorNo==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
  }
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Block Global Vendor");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block Vendor started::"+STIME);
gotoMenu();
gotoVendorSearch();
globalVendor();
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
  
 var VendorNumber = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(VendorNo!=""){
    VendorNumber.HoverMouse();
  VendorNumber.Click();
  WorkspaceUtils.VPWSearchByValue(VendorNumber,"Vendor",VendorNo,"Vendor Number");
    }
    
 var VendorName = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
 VendorName.HoverMouse();
 Sys.HighlightObject(VendorName); 
  VendorName.setText("*");
 
 
 var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
 save.Click();
  TextUtils.writeLog("Company Number, Vendor Number, Currency has entered and Saved in Vendor Search screen");
}

function globalVendor(){ 
  var GblClient = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  waitForObj(GblClient);
  GblClient.HoverMouse();
  GblClient.Click();
  var active = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;  
  waitForObj(active);
  active.HoverMouse();
  active.Click();
  aqUtils.Delay(3000, "Reading from Global Vendor table");
  var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Vendor is available in maconomy to block Global Vendor");
  }  
  aqUtils.Delay(5000, "Playback");
  TextUtils.writeLog("Global Vendor is available in maconomy to block Global Vendor");
}

function goToVendor(){ 
  var home = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
  waitForObj(home);
  Sys.HighlightObject(home);
  home.HoverMouse();
  home.Click();
  
  var information = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.clientpayment;
  waitForObj(information);
  Sys.HighlightObject(information);
  information.HoverMouse(); 
  information.Click();
  ReportUtils.logStep_Screenshot();

  var blockVendor = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget2.Composite.McPopupPickerWidget;
  if(blockVendor.getText()=="Yes")
  ValidationUtils.verify(false,true,"Global Vendor is already blocked");
  else{ 
  blockVendor.Click();
  DropDownList("Yes")
//  blockClient.Keys("Yes");
  var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
  Sys.HighlightObject(save);
  save.HoverMouse(); 
  save.Click();
  ReportUtils.logStep_Screenshot();
  ValidationUtils.verify(true,true,"Global Vendor is Blocked");
 
   if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Vendors - Information")    
    {
    var button = Sys.Process("Maconomy").SWTObject("Shell", "Vendors - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    var label = Sys.Process("Maconomy").SWTObject("Shell", "Vendors - Information").SWTObject("Label", "*").WndCaption;
               button.HoverMouse();
           waitForObj(button);
        Sys.HighlightObject(button);
        button.HoverMouse();
        button.Click();   
                    
     } 
  
  TextUtils.writeLog("Global Vendor is Blocked");
  var AllowForJobs_and_Order = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget3.Composite.McPopupPickerWidget;
  if(AllowForJobs_and_Order.getText()=="No")
  ValidationUtils.verify(true,true,"Allow Payments has Changed to NO");
  else
  ValidationUtils.verify(true,true,"Allow Payments has NOT Changed to NO");
  TextUtils.writeLog("Allow Payments has Changed to NO");
  ReportUtils.logStep_Screenshot();
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

 