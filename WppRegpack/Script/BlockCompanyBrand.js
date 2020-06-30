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
var sheetName = "BlockCompanyBrand";
var ClientNo,BrandNo,Currency,BrandName,brandName ="";
var Language = "";

function Blockcompanybrand(){ 
//  TextUtils.writeLog("Block Gloabl brand Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
  
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("Central Team - Client Account Management","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "BlockCompanyBrand";
ClientNo,BrandNo,Currency ,BrandName,brandName ="";


  ExcelUtils.setExcelName(workBook, "Data Management", true);
  ClientNo = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
  if((ClientNo=="")||(ClientNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  }
  if((ClientNo==null)||(ClientNo=="")){ 
  ValidationUtils.verify(false,true,"Client Number is Needed to Amend Global Client");
  }
  
    ExcelUtils.setExcelName(workBook, "Data Management", true);
  BrandNo = ReadExcelSheet("Global Brand Number",EnvParams.Opco,"Data Management");
  if((BrandNo=="")||(BrandNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
BrandNo = ExcelUtils.getRowDatas("Brand Number",EnvParams.Opco)
  }
if((BrandNo==null)||(BrandNo=="")){ 
ValidationUtils.verify(false,true,"Brand Number is Needed to Amend Global Brand");
}
Log.Message("brandNumber :"+BrandNo)



  ExcelUtils.setExcelName(workBook, "Data Management", true);
  brandName = ReadExcelSheet("Global Brand Name",EnvParams.Opco,"Data Management");
  if((brandName=="")||(brandName==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
brandName = ExcelUtils.getRowDatas("Brand Name",EnvParams.Opco)
  }
if((brandName==null)||(brandName=="")){ 
ValidationUtils.verify(false,true,"Brand Name is Needed to Amend Global Brand");
}
  
  
ExcelUtils.setExcelName(workBook, sheetName, true);
//ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
//  if((ClientNo=="")||(ClientNo==null)){
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  ClientNo = ReadExcelSheet("Client Number",EnvParams.Opco,"Data Management");
//  }
//if((ClientNo==null)||(ClientNo=="")){ 
//ValidationUtils.verify(false,true,"Client Number is Needed to Block Global Brand");
//}
//BrandNo = ExcelUtils.getRowDatas("Brand Number",EnvParams.Opco)
//  if((BrandNo=="")||(BrandNo==null)){
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  BrandNo = ReadExcelSheet("Brand Number",EnvParams.Opco,"Data Management");
//  }
//if((BrandNo==null)||(BrandNo=="")){ 
//ValidationUtils.verify(false,true,"Brand Number is Needed to Block Global Brand");
//}
//
//BrandName = ExcelUtils.getRowDatas("Brand Name",EnvParams.Opco)
//  if((BrandName=="")||(BrandName==null)){
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
//  BrandName = ReadExcelSheet("Brand Name",EnvParams.Opco,"Data Management");
//  }
//if((BrandName==null)||(BrandName=="")){ 
//ValidationUtils.verify(false,true,"Brand Name is Needed to Block Global Brand");
//}

Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Block Global Brand");
}


STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block Company brand started::"+STIME);
gotoMenu();
gotoClientSearch();
globalClient();
client();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
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
TextUtils.writeLog("Entering into Client Management from Accounts Receivable Menu");
}

function gotoClientSearch(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var CompanyNumber = Aliases.ObjectGroup.CompanyNameClientManagement;
 WorkspaceUtils.waitForObj(CompanyNumber)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  CompanyNumber.Click();
  WorkspaceUtils.SearchByValue(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var curr = Aliases.ObjectGroup.CurrencyPicker;
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
 var ClientNumber = Aliases.ObjectGroup.ClientNoField;
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(ClientNo!=""){
  ClientNumber.Click();
  WorkspaceUtils.VPWSearchByValue(ClientNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNo,"Client Number");
    }
    
 var ClientName = Aliases.ObjectGroup.ClientName;
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
 ClientName.setText("*");
 
 
 var save = Aliases.ObjectGroup.SaveButtonClientManagement;
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.RemarksSave;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}


function globalClient(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var GblClient = Aliases.ObjectGroup.JobInfoTab;
// var GblClient = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  GblClient.Click();
  aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//  var active = Aliases.ObjectGroup.ActiveRadioGlobal;
  var active = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//  var table = Aliases.ObjectGroup.GlobalClientTable;
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  //Aliases.ObjectGroup.CompanyClientSearchTable;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
  //  table.getItem(0).
  table.HoverMouse(49, 51);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 51);
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
}



function client(){ 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  //Aliases.Composite3.Composite.PTabFolder.TabFolderPanel.HomeTAB;;
 // NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
 // Aliases.ObjectGroup.HomeTab;
  // Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  home.Click();
  aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  var sublevels = Aliases.ObjectGroup.EmployeeVendorTab;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  sublevels.Click();
  aqUtils.Delay(2000, Indicator.Text);
  TextUtils.writeLog("Navigating to Sub Level");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  var cmpSublevels = Aliases.ObjectGroup.companySublevels;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  cmpSublevels.Click();
  aqUtils.Delay(2000, Indicator.Text);
//  var activeBrand = Aliases.ObjectGroup.ActivecmpBrand;
  var activeBrand = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").
  SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Company Brands").OleValue.toString().trim());
  //Aliases.ObjectGroup.ActiveRadio;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button2;
  activeBrand.Click();
  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Active Brand is selected");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
//  var table =Aliases.ObjectGroup.CompanyClientSearchTable;
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
  // Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//  var brandNmae = Aliases.ObjectGroup.CompanyClientSearchTable.ActiveNameBrand
  var brandNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")
  //Aliases.ObjectGroup.CompanyClientSearchTable.ActiveBrandName;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  brandNmae.Click();
  brandNmae.Keys(brandName);
  aqUtils.Delay(4000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==BrandNo){
//  table.HoverMouse(51, 60);
//  ReportUtils.logStep_Screenshot();
//  table.Click(51, 60);
  
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy to block");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==BrandNo){
//  table.HoverMouse(49, 71);
//  ReportUtils.logStep_Screenshot();  
//  table.Click(49, 71);
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy to block");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==BrandNo){
//  table.HoverMouse(49, 90);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 90);
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy to block");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==BrandNo){
//  table.HoverMouse(49, 109);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 109);
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy to block");
  }
  
  aqUtils.Delay(5000, Indicator.Text);

  TextUtils.writeLog("Company Brand is available in maconomy to block");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  var information = Aliases.ObjectGroup.InformationTab
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  information.Click();
  aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite;
  //Aliases.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.CompanyClientScreen;
  //NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  screen.Click();
  screen.MouseWheel(-200);
  var blockClient = Aliases.ObjectGroup.BlockedIsland;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  if(blockClient.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(false,true,"Company Brand is already blocked");
  else{ 
  blockClient.Click();
  DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
//  blockClient.Keys("Yes");
  aqUtils.Delay(5000, Indicator.Text);
  ReportUtils.logStep_Screenshot();
  var save = Aliases.ObjectGroup.SaveButtonClientManagement;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  save.Click();
  aqUtils.Delay(5000, Indicator.Text);
  ValidationUtils.verify(true,true,"Company Brand is Blocked");
  ReportUtils.logStep_Screenshot();
  TextUtils.writeLog("Company Brand is Blocked");
//  var AllowForJobs_and_Order = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McPopupPickerWidget;
//  if(AllowForJobs_and_Order.getText()=="No")
//  ValidationUtils.verify(true,true,"Allow for use on Jobs and Order has Changed to NO");
//  else
//  ValidationUtils.verify(true,true,"Allow for use on Jobs and Order has NOT Changed to NO");
//  TextUtils.writeLog("Allow for use on Jobs and Order has Changed to NO");
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

 