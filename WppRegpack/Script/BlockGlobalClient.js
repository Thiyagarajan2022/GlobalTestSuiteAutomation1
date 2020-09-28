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
var sheetName = "BlockGlobalClient";
var ClientNo,BrandNo,Currency ="";
var Language = "";
function Blockglobalclient(){ 

Indicator.PushText("waiting for window to open");
aqUtils.Delay(4000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("Central Team - Client Account Management","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "BlockGlobalClient";
ClientNo,BrandNo,Currency ="";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  ClientNo = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
  if((ClientNo=="")||(ClientNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  }
if((ClientNo==null)||(ClientNo=="")){ 
ValidationUtils.verify(false,true,"Client Number is Needed to Amend Global Client");
}
ExcelUtils.setExcelName(workBook, sheetName, true);
Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Block Global Client");
}

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block Gloabl brand started::"+STIME);
gotoMenu();
gotoClientSearch();
globalClient();
client();
aqUtils.Delay(2000, Indicator.Text);
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

ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receivable Menu");
TextUtils.writeLog("Entering into Client Management from Accounts Receivable Menu");
}

function gotoClientSearch(){ 
  
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var CompanyNumber = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  CompanyNumber.Click();
  Log.Message(Language)
  WorkspaceUtils.SearchByValue(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var curr = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
 var ClientNumber = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(ClientNo!=""){
  ClientNumber.Click();
  WorkspaceUtils.VPWSearchByValue(ClientNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNo,"Client Number");
    }
    
 var ClientName = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
 ClientName.setText("*");
 
 
 var save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
 
 TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
}

function globalClient(){ 
  aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var GblClient = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
  GblClient.Click();
  aqUtils.Delay(3000, Indicator.Text);
  var active = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());;
  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var table = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  
  aqUtils.Delay(5000, Indicator.Text);
  TextUtils.writeLog("Global Client is available in maconomy to Amend");
    aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}


function client(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  waitForObj(home);
  Sys.HighlightObject(home);
  home.HoverMouse(); 
  home.Click();
//  aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  waitForObj(information);
  Sys.HighlightObject(information);
  information.HoverMouse(); 
  information.Click();
//  aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
  waitForObj(screen);
  Sys.HighlightObject(screen);
  screen.Click();
  screen.MouseWheel(-200);
  var blockClient = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  waitForObj(blockClient);
  Sys.HighlightObject(blockClient);
  blockClient.HoverMouse(); 
  if(blockClient.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(false,true,"Global Client is already blocked");
  else{
  blockClient.Click();
  DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
//  blockClient.Keys("Yes");
//  aqUtils.Delay(5000, Indicator.Text);
  ReportUtils.logStep_Screenshot();

  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  waitForObj(save);
  Sys.HighlightObject(save);
  save.HoverMouse(); 
  save.Click();
  aqUtils.Delay(6000, "Waiting to Save");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}
aqUtils.Delay(8000, "Saving changes");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}
aqUtils.Delay(8000, "Saving changes");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}
aqUtils.Delay(8000, "Saving changes");
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}
  ValidationUtils.verify(true,true,"Global Client is Blocked");
  ReportUtils.logStep_Screenshot();
  TextUtils.writeLog("Global Client is Blocked");
  var AllowForJobs_and_Order = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McPopupPickerWidget;
  waitForObj(AllowForJobs_and_Order);
  Sys.HighlightObject(AllowForJobs_and_Order);
  AllowForJobs_and_Order.HoverMouse(); 
  if(AllowForJobs_and_Order.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "No").OleValue.toString().trim())
  ValidationUtils.verify(true,true,"Allow for use on Jobs and Order has Changed to NO");
  else
  ValidationUtils.verify(true,true,"Allow for use on Jobs and Order has NOT Changed to NO"); 
  
  TextUtils.writeLog("Allow for use on Jobs and Order has Changed to NO");  
  ExcelUtils.setExcelName(workBook,"Data Management", true);    
  ExcelUtils.WriteExcelSheet("Blocked Global Client No",EnvParams.Opco,"Data Management",ClientNo)
  TextUtils.writeLog("Blocked Global Client No: "+ClientNo);                      
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
            aqUtils.Delay(1000, Indicator.Text);;
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

 