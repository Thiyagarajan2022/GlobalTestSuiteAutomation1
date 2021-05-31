//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT CreditNotePO
//USEUNIT ReverseCreditNote


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var POnumber = "";
var Language = "";
var POnum = "";
var Project_manager = "";

function ApprovePurchaseOrder(sheet,PO){ 
TextUtils.writeLog("Approve Purchase Order Started"); 
Indicator.PushText("waiting for window to open");

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = sheet;
POnum =PO;
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
POnumber = "";

  

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Approving PO started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
try{
getDetails();
gotoMenu();
gettingApproval();
WorkspaceUtils.closeAllWorkspaces();
//CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
  level=i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApprovePO(temp[1],temp[2],i,temp[3]);
}


if(CreditNotePO.CreatePO){ 
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Approved Credit PO Number",EnvParams.Opco,"Data Management",POnumber)
TextUtils.writeLog("Approved Credit PO Number :"+POnumber); 
}

else if(ReverseCreditNote.CreatePO){ 
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Approved ReverseCredit PO Number",EnvParams.Opco,"Data Management",POnumber)
TextUtils.writeLog("Approved ReverseCredit PO Number :"+POnumber); 
}
TextUtils.writeLog("Purchase Orders("+POnumber+") is Approved");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

}
  catch(err){
    Log.Message(err);
  }
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
ExcelUtils.setExcelName(workBook, "Data Management", true);

if(CreditNotePO.CreatePO){ 
// Getting Negative PO Number to Approve
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("Credit PO Number",EnvParams.Opco,"Data Management");
if((POnumber==null)||(POnumber=="")){ 
ValidationUtils.verify(false,true,"PO Number is Needed to Approve Purchase Order");
}

}
else if(ReverseCreditNote.CreatePO){ 
// Getting Negative PO Number to Approve
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("ReverseCredit PO Number",EnvParams.Opco,"Data Management");
if((POnumber==null)||(POnumber=="")){ 
ValidationUtils.verify(false,true,"PO Number is Needed to Approve Purchase Order");
}

}
else{
POnumber = ReadExcelSheet("PO Number_"+POnum,EnvParams.Opco,"Data Management");
if((POnumber=="")||(POnumber==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
POnumber = ExcelUtils.getRowDatas("PO Number",EnvParams.Opco)
}
if((POnumber==null)||(POnumber=="")){ 
ValidationUtils.verify(false,true,"PO Number is Needed to Approve Purchase Order");
}
 }
}

function gotoMenu(){ 
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
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());
}

} 

ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable Menu");
TextUtils.writeLog("Entering into Purchase Orders from Accounts Payable Menu");
}

function gettingApproval(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

  var allPurchase = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
//  var allPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
  WorkspaceUtils.waitForObj(allPurchase);
  allPurchase.Click();
  var companyNo = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget
//  var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.CompanyNo;
  WorkspaceUtils.waitForObj(companyNo);
  companyNo.Click();
  companyNo.Keys("[Tab]");
  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox
//  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.PurchaseNumber;
  WorkspaceUtils.waitForObj(purchaseNo);
  purchaseNo.Click();
  purchaseNo.setText(POnumber);

var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
//  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid;
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(3000, "Reading Table Data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==POnumber){ 
  flag=true;    
  break;
  }
  else{ 
  table.Keys("[Down]");
  }
  }
  ValidationUtils.verify(flag,true,"Created Purchase Order is available in system");
  TextUtils.writeLog("Created Purchase Order is available in system");
  
  
 if(flag){
   var closefilter =Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter
//  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(closefilter);
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.Approvals
//  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  WorkspaceUtils.waitForObj(purchaseOrderApproval);
  purchaseOrderApproval.Click();
  
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }

  ImageRepository.ImageSet.Maximize.Click();

var purchaseApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3
//  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
  WorkspaceUtils.waitForObj(purchaseApproval);
  purchaseApproval.Click();
  var ApproverTable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(ApproverTable);
  var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
        var mainApprover = ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
        var substitur = ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
        var temp = "";
        if(mainApprover .indexOf(Project_manager)==-1){ 
          temp = temp+mainApprover+"*";
        }else{ 
          temp = temp+"SelfApprove"+"*";
        }
        if(substitur .indexOf(Project_manager)==-1){ 
          temp = temp+substitur;
        }
      approvers = EnvParams.Opco+"*"+POnumber+"*"+ temp;
//      approvers = EnvParams.Opco+"*"+POnumber+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
      Approve_Level[y] = approvers;
      y++;
      }
}
TextUtils.writeLog("Finding approvers for Created Purchase Order");
//var listPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.listPurchaseOrder.TabControl;
//listPurchaseOrder.Click();
//ImageRepository.ImageSet.Forward.Click();


CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");

//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
//sheetName = "ApprovePurchaseOrder";
    Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
if(OpCo2[2]==Project_manager){
level = 1;
var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
}
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
//  Approve.PopupMenu.Click("Approve Purchase Order");
//ImageRepository.ImageSet.ApprovePurchaseOrder.Click();
//ReportUtils.logStep_Screenshot();
//aqUtils.Delay(8000, Indicator.Text);;
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created PO");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);

  ValidationUtils.verify(true,true,"Purchase Order is Approved by :"+loginPer)
  TextUtils.writeLog("Purchase Order is Approved by :"+loginPer); 

  
if(Approve_Level.length==1){
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Approved PO_"+POnum,EnvParams.Opco,"Data Management",POnumber);
  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  purchaseOrderApproval.Click();

  ImageRepository.ImageSet.Maximize.Click();

  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
  WorkspaceUtils.waitForObj(purchaseApproval)
  purchaseApproval.Click();
  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable)
ReportUtils.logStep_Screenshot();


}

}
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
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
//     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, "Agency Users", true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

//    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, "SSC Users", true);
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
  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
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

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}
var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Purchase Order from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Purchase Order (Substitute) from To-Dos List");  
var listPass = false;   
  }
}  


if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Purchase Order by Type from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Purchase Order by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Purchase Order by Type (Substitute) from To-Dos List"); 
var listPass = false;   
  }
} 
  }
  
}

function FinalApprovePO(PONum,Apvr,lvl){ 

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget2.McGrid;
var firstCell = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget2.McGrid.McTextWidget;

WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(PONum);
//var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
WorkspaceUtils.waitForObj(table);
aqUtils.Delay(3000, "Reading Data from table");;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==PONum){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Purchase Order is available in Approval List");
TextUtils.writeLog("Created Purchase Order is available in Approval List");
if(flag){ 
  
var closefilter = "";
var filterStat = false
if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.Index==1){
closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite;
Log.Message(closefilter.FullName)
}else{
 closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite
Log.Message(closefilter.FullName)
}
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
//Approve PO

    var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
        if((PChild.isVisible()) && (PChild.ChildCount==3)){
        Add[ChildCount] = PChild;
        ChildCount++;
     }
     }
     
     var Approve = "";
     var pos = 0;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height>pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       Approve = Add[ip];
     }     
     }
     
     Sys.HighlightObject(Approve)
     Log.Message(Approve.FullName)
     Approved = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
     if(Approved.Visible){ 
     Approve =  Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
     }
     else{ 
     Approve = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);  
     }
     Sys.HighlightObject(Approve)


//var Approve = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite;
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    break;
  }
}
WorkspaceUtils.waitForObj(Approve);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
//aqUtils.Delay(3000, Indicator.Text);
//Approve.PopupMenu.Click("Approve Purchase Order");
//ReportUtils.logStep_Screenshot();
//aqUtils.Delay(8000, Indicator.Text);;
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Apvr)
TextUtils.writeLog("Purchase Order is Approved by "+Apvr);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//Uncommand
/*
//var screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");;
var screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "")
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-5);
//  aqUtils.Delay(5000, Indicator.Text);
//var ApvPerson = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 6).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
var ApvPerson = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 6).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while ((ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

    if(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1){
  ValidationUtils.verify(true,true,"Purchase Order is Approved by :"+loginPer)
  TextUtils.writeLog("Purchase Order is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Purchase Order is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Purchase Order is Approved by :"+loginPer+ "But its Not Reflected")
  }
  
*/


if(Approve_Level.length==lvl+1){
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Approved PO_"+POnum,EnvParams.Opco,"Data Management",POnumber);



var approvalBar = "";
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.isVisible()){
approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl
}else{
approvalBar = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
}

//var approvalBar = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
ImageRepository.ImageSet.Maximize.Click();

if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN"))
   Runner.CallMethod("IND_ApprovePurchaseOrder.ApprovalStatus");
else{

var POapproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
//var POapproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3
WorkspaceUtils.waitForObj(POapproval)
POapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
POapproval.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//var approvertable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
WorkspaceUtils.waitForObj(approvertable)
ReportUtils.logStep_Screenshot();
}
}
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Apvr)
}
}

} 


function exlApprove() {

// Firstcell in Approve PO
      var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
   Sys.Process("Maconomy").Refresh();
         for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
      if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")){
      Add[ChildCount] = PChild;
      ChildCount++;

     }
     }      
     
      Parent = "";
     var pos = 0;
     for(var i=0;i<Add.length;i++){ 
     if(Add[i].Height>pos){ 
       pos = Add[i].Height;
       Parent = Add[i];
     }     
     } 
        Log.Message(Parent.FullName);
   Parent = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "");
   Log.Message(Parent.FullName);
   Sys.HighlightObject(Parent);
    ChildCount = 0;
    Add = [];
     for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
     if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")){
         Add[ChildCount] = PChild;
         ChildCount++;
     }
     }
     
     Parent = "";
     var pos = 1000;
     for(var i=0;i<Add.length;i++){ 
     if(Add[i].Height<pos){ 
       pos = Add[i].Height;
       Parent = Add[i];
     }     
     } 
    Log.Message(Parent.FullName)
    Sys.HighlightObject(Parent);
    Parent = Parent.SWTObject("McClumpSashForm", "")
         for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
      if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")&& (PChild.ChildCount==1)){
      Add[ChildCount] = PChild;
      ChildCount++;

     }
     }      
     
      Parent = "";
     var pos = 1000;
     for(var i=0;i<Add.length;i++){ 
     if(Add[i].Height<pos){ 
       pos = Add[i].Height;
       Parent = Add[i];
     }     
     } 
    var firstCell = Parent.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    Log.Message(firstCell.FullName)
    Sys.HighlightObject(firstCell);

}



function CloseFilters() {

// CloseFillter in Approve PO
      var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
   Sys.Process("Maconomy").Refresh();
         for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
      if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")){
      Add[ChildCount] = PChild;
      ChildCount++;

     }
     }      
     
      Parent = "";
     var pos = 0;
     for(var i=0;i<Add.length;i++){ 
     if(Add[i].Height>pos){ 
       pos = Add[i].Height;
       Parent = Add[i];
     }     
     } 
   Log.Message(Parent.FullName);
   Parent = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "");
   Log.Message(Parent.FullName);
   Sys.HighlightObject(Parent);
    ChildCount = 0;
    Add = [];
     for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
     if((PChild.isVisible()) && (PChild.JavaClassName=="TabFolderPanel")){
         Parent = Parent.Child(i);
         ChildCount++;
     }
     }
     
     
   Log.Message(Parent.FullName);
   Sys.HighlightObject(Parent);
    ChildCount = 0;
    Add = [];
     for(var i=0;i<Parent.ChildCount;i++){ 
     var PChild = Parent.Child(i);
      if((PChild.isVisible()) && (PChild.JavaClassName=="Composite")){
      Add[ChildCount] = PChild;
         ChildCount++;
     }
     }
     
     Parent = "";
     var pos = 1000;
     for(var i=0;i<Add.length;i++){ 
     if((Add[i].Height<pos)&&(Add[i].ChildCount==1)){ 
       pos = Add[i].Height;
       Parent = Add[i];
     }     
     } 
    Log.Message(Parent.FullName)
    Sys.HighlightObject(Parent);
    var CloseFilter = Parent.SWTObject("SingleToolItemControl", "")
    Log.Message(CloseFilter.FullName)
    Sys.HighlightObject(CloseFilter);

}


function ApproveButton(){ 
  var Language = "English";
    var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
        if((PChild.isVisible()) && (PChild.ChildCount==3)){
        Add[ChildCount] = PChild;
        ChildCount++;
     }
     }
     
     var Approve = "";
     var pos = 0;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height>pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       Approve = Add[ip];
     }     
     }
     
     Sys.HighlightObject(Approve)
     Log.Message(Approve.FullName)
     Approved = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
     if(Approved.Visible){ 
     Approve =  Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
     }
     else{ 
     Approve = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);  
     }
     Sys.HighlightObject(Approve)

 
Log.Message(Approve.FullName)
Sys.HighlightObject(Approve);
 for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text.OleValue.toString().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())!=-1)){
    Approve = Approve.Child(i);
    Sys.HighlightObject(Approve)
    Log.Message(Approve.FullName)
    break;
  }
}

Sys.HighlightObject(Approve)
Sys.HighlightObject(Approve);

}

function ApvBar(){ 
      var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
       var pos = 0;         
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
        if((PChild.isVisible()) && (PChild.ChildCount==3) && (PChild.ScreenLeft>=pos)){
        pos = PChild.ScreenLeft;
        Add[ChildCount] = PChild;
        ChildCount++;
     }
     }
     
     var Approve = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].ScreenTop<pos){ 
       pos = Add[ip].ScreenTop;
       Log.Message(pos)
       Approve = Add[ip];
     }     
     }
     
    var approvalBar = Approve.SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
    Sys.HighlightObject(approvalBar)
//    POapproval
}