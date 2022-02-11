//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "BlockCompanyClient";
var ClientNo,BrandNo,Currency ="";

function BlockCompanyclient(){ 
//  TextUtils.writeLog("Block Company Client Started"); 
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("Central Team - Client Account Management","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "BlockCompanyClient";
ClientNo,BrandNo,Currency ="";

 ExcelUtils.setExcelName(workBook, "Data Management", true);
  ClientNo = ReadExcelSheet("Company Client Number",EnvParams.Opco,"Data Management");
  if((ClientNo=="")||(ClientNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  }
  if((ClientNo==null)||(ClientNo=="")){ 
  ValidationUtils.verify(false,true,"Client Number is Needed to Amend Global Client");
  }
    Log.Message("ClientNumber"+ClientNo)


ExcelUtils.setExcelName(workBook, sheetName, true);

Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Block Global Client");
}



STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block Global client started::"+STIME);
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
//TextUtils.writeLog("Entering into Purchase Orders from Accounts Payable Menu");
}


function gotoClientSearch(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var CompanyNumber =Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
CompanyNumber.Click();
WorkspaceUtils.SearchByValue(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
aqUtils.Delay(5000, Indicator.Text);
var curr = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Currency;
curr.Keys(" ");
aqUtils.Delay(5000, Indicator.Text);
if(Currency!=""){
curr.Click();
WorkspaceUtils.DropDownList(Currency,"Currency")
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var ClientNumber = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.ClientNoFiled;
if(ClientNo!=""){
ClientNumber.Click();
WorkspaceUtils.VPWSearchByValue(ClientNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNo,"Client Number");
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}   
var ClientName = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.ClientName
ClientName.setText("*");
 
 
var save = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SaveButton;
save.Click();
aqUtils.Delay(5000, Indicator.Text);
}



function globalClient(){ 
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var GblClient = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.CompanyClientTab;
  GblClient.Click();
  aqUtils.Delay(3000, Indicator.Text);
  var active = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());;
  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var table = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.CompanyClientTable;
  aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
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
    aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}


function test()
{
   var table = Aliases.ObjectGroup.CompanyClientSearchTable;
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
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var home =Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Home;
  
 aqUtils.Delay(2000, Indicator.Text);
 home.Click();
  aqUtils.Delay(2000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var information =Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.InformationTab

  information.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  aqUtils.Delay(2000, Indicator.Text);
  var screen =Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite
  //Aliases.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.CompanyClientScreen;
  //NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  screen.Click();
  screen.MouseWheel(-200);
  aqUtils.Delay(2000, Indicator.Text);
  var blockClient = Aliases.ObjectGroup.BlockedIsland;
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  if(blockClient.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(false,true,"Company Client is already blocked");
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
  ValidationUtils.verify(true,true,"Company Client is Blocked");
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

