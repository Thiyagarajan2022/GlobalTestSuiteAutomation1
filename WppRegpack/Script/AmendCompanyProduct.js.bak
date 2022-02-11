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
var sheetName = "AmendCompanyProduct";
var ClientNo,ProductNo,Currency,ProductName,Phone,Email ="";


function AmendCompanyproduct(){ 
//  TextUtils.writeLog("Block Company Client Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
//var Project_manager = EnvParams.Opco+" Finance";
Project_manager = ExcelUtils.getRowDatas("Central Team - Client Account Management","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "AmendCompanyProduct";
ClientNo,ProductNo,Currency,ProductName,Phone,Email ="";



STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block Global client started::"+STIME);
getDetails();
gotoMenu();
gotoClientSearch();
globalClient();
client();
menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

}

function getDetails(){ 
  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  ClientNo = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
  ClientNo = ReadExcelSheet("Company Client Number",EnvParams.Opco,"Data Management");

  if((ClientNo=="")||(ClientNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  }
if((ClientNo==null)||(ClientNo=="")){ 
ValidationUtils.verify(false,true,"Client Number is Needed to Block Global Client");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Block Global Client");
}

  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  ProductNo = ReadExcelSheet("Global Product Number",EnvParams.Opco,"Data Management");
  ProductNo = ReadExcelSheet("Company Product Number",EnvParams.Opco,"Data Management");
  if((ProductNo=="")||(ProductNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
ProductNo = ExcelUtils.getRowDatas("Product Number",EnvParams.Opco)
  }
if((ProductNo==null)||(ProductNo=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Block Global Product");
}

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  ProductName = ReadExcelSheet("Company Product Name",EnvParams.Opco,"Data Management");
  if((ProductName=="")||(ProductName==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
ProductName = ExcelUtils.getRowDatas("Product Name",EnvParams.Opco)
  }
if((ProductName==null)||(ProductName=="")){ 
ValidationUtils.verify(false,true,"Product Name is Needed to Block Global Product");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
//Add1 = ExcelUtils.getRowDatas("Address_1",EnvParams.Opco)
//Add2 = ExcelUtils.getRowDatas("Address_2",EnvParams.Opco)
//Add3 = ExcelUtils.getRowDatas("Address_3",EnvParams.Opco)
Phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
Email = ExcelUtils.getRowDatas("E-mail",EnvParams.Opco)
if(((Phone==null)||(Phone==""))&&((Email==null)||(Email==""))){ 
ValidationUtils.verify(false,true,"Phone,Email is Needed to Amend Global Client");
}

}


function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Account_Receivable.Exists()){
ImageRepository.ImageSet.Account_Receivable.Click();// GL
}
else if(ImageRepository.ImageSet.Acc_Receivable_1.Exists()){
ImageRepository.ImageSet.Acc_Receivable_1.Click();
}
else{
ImageRepository.ImageSet.Acc_Receivable_2.Click();
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim());
}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receivable Menu");
//TextUtils.writeLog("Entering into Purchase Orders from Accounts Payable Menu");
}

function gotoClientSearch(){ 
  Delay(1000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var CompanyNumber = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  CompanyNumber.Click();
  Delay(3000);
  WorkspaceUtils.SearchByValue(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");

 var curr = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  
var ClientNumber = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(ClientNo!=""){
  ClientNumber.Click();
  WorkspaceUtils.VPWSearchByValue(ClientNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNo,"Client Number");
    }
    
 var ClientName =  Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;

ClientName.setText("*");
 
 
 var save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
}

function globalClient(){ 
 // var GblClient = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
 // GblClient.Click();
  aqUtils.Delay(3000, Indicator.Text);
  aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var active = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
   var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table)
  if(table.getItem(0).getText(0).OleValue.toString().trim()==ClientNo){
  //  table.getItem(0).
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  
  aqUtils.Delay(5000, Indicator.Text);
   TextUtils.writeLog("Global Client is available in maconomy to block Global Product");
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}
function test()
{
   var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
   table.HoverMouse(51, 60);
if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
  //  table.getItem(0).
  table.HoverMouse(51, 60);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 60);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
}


function client(){ 
  aqUtils.Delay(10000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var home = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
 // Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(home); 
home.Click();
  aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//  var sublevels = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
    var ChildCount = 0;
    var Add = [];

   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
for(var ip=0;ip<Parent.ChildCount;ip++){ 
var PChild = Parent.Child(ip);
if((PChild.isVisible()) && (PChild.ChildCount==1)){
Add[ChildCount] = PChild;
ChildCount++;
}
}
     
     var Approve = "";
     var sublevels= ""
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       Approve = Add[ip];
     }     
     }
     
     Sys.HighlightObject(Approve)
     Log.Message(Approve.FullName)
     sublevels = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);

     Sys.HighlightObject(sublevels)
Sys.HighlightObject(sublevels);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  sublevels.Click();
  
aqUtils.Delay(2000, Indicator.Text);
  TextUtils.writeLog("Navigating to Sub Level");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var cpySublevels = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  cpySublevels.Click();
  aqUtils.Delay(2000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var activeProduct = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
  activeProduct.Click();
  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Active Product is selected");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table);

//var ProductNmae = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "");
////Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
////Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget  
//Sys.HighlightObject(ProductNmae);
//  ProductNmae.Click();
//  ProductNmae.Keys(ProductName);
  aqUtils.Delay(4000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  //Log.Message(table.getItem(0).getText_2(0).OleValue.toString().trim());
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ProductNo){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy to block");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ProductNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy to block");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ProductNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy to block");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ProductNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy to block");
  }
  
  aqUtils.Delay(5000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Global Product is available in maconomy to block");
  var home=Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl3;
  //Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl3;
Sys.HighlightObject(home);  
home.Click();

  var information = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
  //Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl3;
 // Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl3;
 Sys.HighlightObject(information);
 information.Click();
  
aqUtils.Delay(2000, Indicator.Text);
  //var screen = Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10  
  //screen.Click();
  //screen.MouseWheel(-200);

//  var C_Phone = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
  var C_Phone = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.PhoneNo
  
Sys.HighlightObject(C_Phone);
  //Aliases.ObjectGroup.CompanyClientPhoneNo;
  
//  var C_mail = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget;
  var C_mail = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Mail
 
 Sys.HighlightObject(C_mail);
 //Aliases.ObjectGroup.CompanyClientEmail;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite9.McTextWidget;
//  var Address1 = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.SWTObject("McTextWidget", "", 2);
//                 
//  var Address2 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.SWTObject("McTextWidget", "", 2);
//  var Address3 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  
  var changes = false;
//      if(Add1!=""){
//    if(Address1.getText()!=Add1){
//       Sys.HighlightObject(Address1);
//      Address1.Click();
//    Address1.setText(Add1);
//    ValidationUtils.verify(true,true,"Add1 Number is Changed");
//    changes = true;
//    }
//    else
//    ReportUtils.logStep("INFO","Given Add1 in datasheet is as same as Value in Maconomy")
//    }
//     
//      if(Add2!=""){
//    if(Address2.getText()!=Add2){
//       Sys.HighlightObject(Address2);
//      Address2.Click();
//    Address2.setText(Add2);
//    ValidationUtils.verify(true,true,"Add2 Number is Changed");
//    changes = true;
//    }
//    else
//    ReportUtils.logStep("INFO","Given Add2 in datasheet is as same as Value in Maconomy")
//    }
//    
//    
//      if(Add3!=""){
//    if(Address3.getText()!=Add3){
//       Sys.HighlightObject(Address3);
//      Address3.Click();
//    Address3.setText(Add3);
//    ValidationUtils.verify(true,true,"Add3 Number is Changed");
//    changes = true;
//    }
//    else
//    ReportUtils.logStep("INFO","Given Add3 in datasheet is as same as Value in Maconomy")
//    }
    
    
    if(Phone!=""){
    if(C_Phone.getText()!=Phone){
       Sys.HighlightObject(C_Phone);
      C_Phone.Click();
    C_Phone.setText(Phone);
    ValidationUtils.verify(true,true,"Phone Number is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Phone Number in datasheet is as same as Value in Maconomy")
    }
    

    if(Email!=""){
    if(C_mail.getText()!=Email){
      Sys.HighlightObject(C_mail);
      C_mail.Click();
    C_mail.setText(Email);
    ValidationUtils.verify(true,true,"Email is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Email in datasheet is as same as Value in Maconomy")
    }
    
        if(changes){ 
    
  var save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  //Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;  
//Aliases.Maconomy.Screen4.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//Aliases.ObjectGroup.SaveButtonClientManagement;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  save.Click();
  aqUtils.Delay(5000, Indicator.Text);
  ValidationUtils.verify(true,true,"Company Client is Amended");
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

