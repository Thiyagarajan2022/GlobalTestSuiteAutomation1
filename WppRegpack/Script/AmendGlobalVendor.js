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
var sheetName = "AmendGlobalVendor";
var VendorNo,email,VendorName,PhoneNum,TaxNum,CompanyReg,Currency,Streat1 ="";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var Language = "";

function AmendGlobalVendor(){ 
TextUtils.writeLog("Amend Gloabl Vendor Started");
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);

Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("Central Team - Vendor Account Management","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "AmendGlobalVendor";
VendorNo,email,VendorName,Streat1,PhoneNum ="";
ExcelUtils.setExcelName(workBook, sheetName, true);
Currency = ExcelUtils.getRowDatas("currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Amend Global Vendor");
}

Streat1 = ExcelUtils.getRowDatas("Street 1",EnvParams.Opco)
//if((PhoneNum==null)||(PhoneNum=="")){ 
//ValidationUtils.verify(false,true,"Phone Number is Needed to Amend Global Vendor");
//}
email = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
//if((email==null)||(email=="")){ 
//ValidationUtils.verify(false,true,"Email is Needed to Amend Global Vendor");
//}

PhoneNum = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
//if((PhoneNum==null)||(PhoneNum=="")){ 
//ValidationUtils.verify(false,true,"Phone Number is Needed to Amend Global Vendor");
//}
TaxNum = ExcelUtils.getRowDatas("TaxNo",EnvParams.Opco)
//if((TaxNum==null)||(TaxNum=="")){ 
//ValidationUtils.verify(false,true,"Tax Number is Needed to Amend Global Vendor");
//}
CompanyReg = ExcelUtils.getRowDatas("CompRegNo",EnvParams.Opco)
//if((CompanyReg==null)||(CompanyReg=="")){ 
//ValidationUtils.verify(false,true,"Company Reg Number is Needed to Amend Global Vendor");
//}
VendorName = ExcelUtils.getRowDatas("Vendor Name",EnvParams.Opco)
//if((VendorName==null)||(VendorName=="")){ 
//ValidationUtils.verify(false,true,"Vendor Name is Needed to Amend Global Vendor");
//}

ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Amend Global Vendor");
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
globalVendor();
goToVendor();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
    for(var i=level;i<ApproveInfo.length;i++){
      level=i;
      WorkspaceUtils.closeMaconomy();
      aqUtils.Delay(10000, Indicator.Text);
      var temp = ApproveInfo[i].split("*");
      Restart.login(temp[2]);
      aqUtils.Delay(5000, Indicator.Text);
      todo(temp[3]);
      FinalApprove(temp[1],temp[2]);
    }
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Management").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor Management").OleValue.toString().trim());
}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Vendor Management from Accounts Payable Menu");
}

function gotoVendorSearch(){ 
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);

 var CompanyNumber = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget; 
  waitForObj(CompanyNumber);
  Sys.HighlightObject(CompanyNumber)
  CompanyNumber.Click();
  WorkspaceUtils.SearchByValue(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

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
  WorkspaceUtils.VPWSearchByValue(VendorNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
    }
    
 var Vendor_Name = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  Vendor_Name.HoverMouse();
 Sys.HighlightObject(Vendor_Name);  
  Vendor_Name.setText("*");
 
 var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
 save.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  TextUtils.writeLog("Company Number, Vendor Number, Currency has entered and Saved in Vendor Search screen");
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);

}

function globalVendor(){ 
    
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);

  var GblClient = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  waitForObj(GblClient);
  GblClient.HoverMouse();
  GblClient.Click();
  var active = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Vendors").OleValue.toString().trim());;; 
   waitForObj(active);
  active.HoverMouse();
  Sys.HighlightObject(active);
  active.Click();
  aqUtils.Delay(3000, "Reading from Global Vendor table");
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);

  var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Amend Global Vendor is available in maconomy");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Amend Global Vendor is available in maconomy");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Amend Global Vendor is available in maconomy");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==VendorNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Amend Global Vendor is available in maconomy");
  }  
  aqUtils.Delay(5000, "Playback");
  TextUtils.writeLog("Amend Global Vendor is available in maconomy");
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);

}

function goToVendor(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);

  var home = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
  Sys.HighlightObject(home);
  home.HoverMouse();   
  home.Click();

  TextUtils.writeLog("Amend Global Vendor is available in maconomy to change information window");  
  var information = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.clientpayment;
  waitForObj(information);
  Sys.HighlightObject(information);
  information.HoverMouse(); 
  information.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(10000, Indicator.Text);
  
  var Vendor_Name = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  var phone =       NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  var Tax_No = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
   var Comp_Reg_No = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  var currency = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite2.McPopupPickerWidget;
  var Str1 = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2)
  var Ven_Email = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 12).SWTObject("McTextWidget", "", 2);
  var Ven_reEmail = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 13).SWTObject("McTextWidget", "", 2);
  
Sys.HighlightObject(Str1);
Sys.HighlightObject(phone);
Sys.HighlightObject(Ven_Email);
Sys.HighlightObject(Str1);
Sys.HighlightObject(Ven_reEmail);
  var changes = false;
 
//   if(VendorName!=""){
//    if(Vendor_Name.getText()!=VendorName){
//    Vendor_Name.setText(VendorName);
//    ValidationUtils.verify(true,true,"Vendor Name is Changed");
//    TextUtils.writeLog("Vendor Name is Changed");
//    changes = true;
//    }
//    else
//    ReportUtils.logStep("INFO","Given Vendor Name in datasheet is as same as Value in Maconomy")
//    }
     if(PhoneNum!=""){
    if(phone.getText()!=PhoneNum){
    phone.setText(PhoneNum);
    ValidationUtils.verify(true,true,"Phone Number is Changed");
    TextUtils.writeLog("Phone Number is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Phone Number in datasheet is as same as Value in Maconomy")
    }    
    
    if(Streat1!=""){
    if(Str1.getText()!=Streat1){
    Str1.setText(Streat1);
    ValidationUtils.verify(true,true,"Streat1 is Changed");
    TextUtils.writeLog("Streat1 is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Streat1 in datasheet is as same as Value in Maconomy")
    }   
    
    if(email!=""){
    if(Ven_Email.getText()!=email){
    Ven_Email.setText(email);
    ValidationUtils.verify(true,true,"Email is Changed");
    TextUtils.writeLog("Email is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Email in datasheet is as same as Value in Maconomy")
    }  
    
    if(email!=""){
    if(Ven_reEmail.getText()!=email){
    Ven_reEmail.setText(email);
    ValidationUtils.verify(true,true,"Remittance Email is Changed");
    TextUtils.writeLog("Remittance Email is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Remittance Email in datasheet is as same as Value in Maconomy")
    }  
    
    
//      if(TaxNum!=""){
//    if(Tax_No.getText()!=TaxNum){
//    Tax_No.setText(TaxNum);
//    ValidationUtils.verify(true,true,"Tax Number is Changed");
//    TextUtils.writeLog("Tax Number is Changed");
//    changes = true;
//    }
//    else
//    ReportUtils.logStep("INFO","Given Tax Number in datasheet is as same as Value in Maconomy")
//    }   
     
//    if(CompanyReg!=""){
//    if(Comp_Reg_No.getText()!=CompanyReg){
//    Comp_Reg_No.setText(CompanyReg);
//    ValidationUtils.verify(true,true,"CompanyReg Number is Changed");
//    TextUtils.writeLog("CompanyReg Number is Changed");
//    changes = true;
//    }
//    else
//    ReportUtils.logStep("INFO","Given CompanyReg Number in datasheet is as same as Value in Maconomy")
//    }   
    
    aqUtils.Delay(3000,"Playback");
//    if(changes){ 
  var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
    waitForObj(save);
  Sys.HighlightObject(save);
  save.HoverMouse();    
  save.Click();     
    aqUtils.Delay(10000,"Playback");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
    
     var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendors - Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendors - Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
}

     
  
  ReportUtils.logStep_Screenshot();
  ValidationUtils.verify(true,true,"Amend Global Vendor field are updated and saved in macanomy");
  TextUtils.writeLog("Amend Global Vendor field are updated and saved in macanomy");
      aqUtils.Delay(10000,"Playback");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
      var submit = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      waitForObj(submit);
      Sys.HighlightObject(submit);
      submit.Click();
      ValidationUtils.verify(true,true,"Amend Global Vendor updated and Submitted in macanomy");
      aqUtils.Delay(10000,"Playback");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
    var printstatus = false;
    if(!printstatus)    
    if((NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6).isVisible())
    {
    var aaprover_pane = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.TabControl;
    printstatus = true;
    }
//     if(!printstatus)    
//    if((NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6).isVisible())
//    {
//    var aaprover_pane = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.TabControl;
//    printstatus = true;
//    }
     aaprover_pane.Click();
     if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
       ImageRepository.ImageSet0.Maximize.Click();  
     if(ImageRepository.ImageSet.Tab_Icon.Exists()){}      
           var printstatus = false;
    if(!printstatus)    
    if((NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6).isVisible())
    {
    var VendorApproval = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;       
      printstatus = true;
    }
    if((NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2).isVisible())
    {
    var VendorApproval = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;       
      printstatus = true;
    }
     VendorApproval.Click();
     
     if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
     
      var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
       Sys.HighlightObject(ApproverTable);
         var y=0;
                 for(var i=0;i<ApproverTable.getItemCount();i++){   
                 var approvers="";
                  if(ApproverTable.getItem(i).getText_2(5)!="Approved"){
                  approvers = EnvParams.Opco+"*"+VendorNo+"*"+ApproverTable.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
                  Log.Message("Approver level :" +i+ ": " +approvers);
                  Approve_Level[y] = approvers;
                  Log.Message(Approve_Level[y])
                  y++;
                  }
              }
                 
    var CloseBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
    CloseBar.Click();
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    ImageRepository.ImageSet0.Forward.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
//var OpCo1 = EnvParams.Opco;
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
sheetName = "AmendGlobalClient";
if(OpCo2[2]==Project_manager){
    level = 1;
    
    var Approve = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
    Sys.HighlightObject(Approve)
    if(Approve.isEnabled()){ 
    Approve.HoverMouse();
    ReportUtils.logStep_Screenshot();
      Approve.Click(); 
    ReportUtils.logStep_Screenshot();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(8000, Indicator.Text);;
    ValidationUtils.verify(true,true,"Amend Global Vendor is Approved by "+Project_manager)
    TextUtils.writeLog("Levels 0 has  Approved Amend Global Vendor");
    aqUtils.Delay(8000, Indicator.Text);;
    }
    }    
//      }      
//    }else{ 
//      ValidationUtils.verify(false,true,"There is no changes happen in Maconomy screen")
//    }
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
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("UAE")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("UAE")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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
aqUtils.Delay(5000, Indicator.Text); 
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
          Log.Message(Client_Managt.getItem(j).getText().OleValue.toString().trim())
          Log.Message(temp +" : "+temp.indexOf((JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2))
          if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
            Client_Managt.ClickItem("|"+temp);   
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp);  
            TextUtils.writeLog("Entering into Approve Vendor by Type from To-Dos List");
            listPass = false; 
            break;
          }
      }
      if(lvl==3)
      for(var j=0;j<Client_Managt.getItemCount();j++){ 
          var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
          var temp1 = temp.split("(");
          Log.Message(temp+" : "+temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)
        if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor by Type (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
            Client_Managt.ClickItem("|"+temp);    
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp); 
            TextUtils.writeLog("Entering into Approve Vendor by Type (Substitute) from To-Dos List");
            var listPass = false;   
            break;
         }
      }  
  if(listPass){
    if(lvl==2)
          for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
            Log.Message(temp+" : "+(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2))
              if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
                  Client_Managt.ClickItem("|"+temp);   
                  ReportUtils.logStep_Screenshot(); 
                  Client_Managt.DblClickItem("|"+temp);  
                  TextUtils.writeLog("Entering into Approve Vendor from To-Dos List");
                  listPass = false; 
                  break;
                }
           }
    if(lvl==3)
        for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
            Log.Message(temp+" : "+temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor (Substitute)").OleValue.toString().trim()+" (")!=-1)
          if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Vendor (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
              Client_Managt.ClickItem("|"+temp);    
              ReportUtils.logStep_Screenshot(); 
              Client_Managt.DblClickItem("|"+temp); 
              TextUtils.writeLog("Entering into Approve Vendor (Substitute) from To-Dos List");
              var listPass = true;   
              break;
            }
        } 
  }
}


function FinalApprove(C_Num,Apvr){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(10000, "Reading Data in table");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(10000, "Reading Data in table");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(10000, "Reading Data in table");;

var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
var showFilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
waitForObj(table);
Sys.HighlightObject(showFilter);
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.Click();
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(10000, "Reading Data in table");;

var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
firstCell.setText(C_Num);
var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  
aqUtils.Delay(6000, Indicator.Text);;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==C_Num){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Amended Vendor is available in Approval List");
TextUtils.writeLog("Amended Vendor is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(10000, "Reading Data in table");;

var Approve = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
TextUtils.writeLog("Approve button is Clicked"); 
aqUtils.Delay(9000, "Waiting to Approve");;
ValidationUtils.verify(true,true,"Amended Vendor is Approved by "+Apvr)
TextUtils.writeLog("Amended Vendor is Approved by "+Apvr);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(10000, "Reading Data in table");;

}
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

 