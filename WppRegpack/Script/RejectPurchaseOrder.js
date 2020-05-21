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
var sheetName = "RejectPurchaseOrder";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var POnumber = ""
function RejectPurchaseOrder(){ 
TextUtils.writeLog("Reject Purchase Order Started");
Indicator.PushText("waiting for window to reponse");
//aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = EnvParams.Opco+" Finance";
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
sheetName = "RejectPurchaseOrder";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
POnumber = "";
//VendorID,Job_Number,WorkCode,Detailed_Description,Qly,UnitPrice = "";
  
Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Approving PO started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME);



getDetails();
gotoMenu();
gettingApproval();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApprovePO(temp[1],temp[2],i);
break;
}
TextUtils.writeLog("Purchase Orders("+POnumber+") is Rejected");
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
ExcelUtils.setExcelName(workBook, "Data Management", true);
POnumber = ReadExcelSheet("PO Number",EnvParams.Opco,"Data Management");
  if((POnumber=="")||(POnumber==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
POnumber = ExcelUtils.getRowDatas("PO Number",EnvParams.Opco)
  }
if((POnumber==null)||(POnumber=="")){ 
ValidationUtils.verify(false,true,"PO Number is Needed to Approve Purchase Order");
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
//aqUtils.Delay(3000, Indicator.Text);
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
Client_Managt.ClickItem("|Purchase Orders");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Purchase Orders");
TextUtils.writeLog("Entering into Purchase Orders from Accounts Payable Menu");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable Menu");

}

function gettingApproval(){ 
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

  var allPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.AllPurchaseOrder;
  WorkspaceUtils.waitForObj(allPurchase);
  allPurchase.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.CompanyNo;
  WorkspaceUtils.waitForObj(companyNo);
  companyNo.Click();
  companyNo.Keys("[Tab]");
  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.PurchaseNumber;
  WorkspaceUtils.waitForObj(purchaseNo);
  purchaseNo.Click();
  purchaseNo.setText(POnumber);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid;
  WorkspaceUtils.waitForObj(table);
  aqUtils.Delay(3000, "Reading Table Data");
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
  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(closefilter);
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  WorkspaceUtils.waitForObj(purchaseOrderApproval);
  purchaseOrderApproval.Click();
//  if(ImageRepository.ImageSet.Forward.Exists()){ 
//   if(ImageRepository.ImageSet.Maximize.Exists()){
//   ImageRepository.ImageSet.Maximize.Click();
//   }
//  }else{
//    var approveAction = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
//    WorkspaceUtils.waitForObj(approveAction);
//    approveAction.HoverMouse();
//    ReportUtils.logStep_Screenshot();
//    approveAction.Click();
//    aqUtils.Delay(3000, Indicator.Text);;
//    if(ImageRepository.ImageSet.Maximize.Exists()){
    ImageRepository.ImageSet.Maximize.Click();
//    }
//  }
//  aqUtils.Delay(3000, Indicator.Text);;
  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
  WorkspaceUtils.waitForObj(purchaseApproval);
  purchaseApproval.Click();
  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(ApproverTable);
   var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(6)!="Approved"){
      approvers = EnvParams.Opco+"*"+POnumber+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
//      Approve_Level[y] = "1307*1307100026*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
      Approve_Level[y] = approvers;
      y++;
      }
}
TextUtils.writeLog("Finding approvers for Created Purchase Order");
CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
//var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
sheetName = "RejectPurchaseOrder";
if(OpCo2[2]==Project_manager){
//var OpCo1 = EnvParams.Opco;
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
//var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
//if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
//if((Project_manager.indexOf(Approve_Level[0])!=-1)||(Project_manager.indexOf(OpCo2)!=-1)){
level = 1;
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
var remarks = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid.Remarks;
remarks.setText("Rejected");
aqUtils.Delay(1000, Indicator.Text);;
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.RemarksSave;
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
//ImageRepository.ImageSet.Rejected.HoverMouse();
ImageRepository.ImageSet.Rejected.Click();
ReportUtils.logStep_Screenshot();
//aqUtils.Delay(8000, Indicator.Text);;
aqUtils.Delay(1000, Indicator.Text);;
ValidationUtils.verify(true,true,"Purchase Order is Rejected by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Rejected the Purchase Order by "+Project_manager);

var ii=0;
while ((approvertable.getItem(0).getText_2(6).OleValue.toString().trim()!="Rejected")&&(ii!=60))
{
  aqUtils.Delay(100);
  ii++;
  approvertable.Refresh();
}
//Log.Message(approvertable.getItem(0).getText_2(3).OleValue.toString().trim())
//Log.Message(approvertable.getItem(0).getText_2(4).OleValue.toString().trim())
//Log.Message(approvertable.getItem(0).getText_2(5).OleValue.toString().trim())
//Log.Message(approvertable.getItem(0).getText_2(6).OleValue.toString().trim())
//Log.Message(approvertable.getItem(0).getText_2(7).OleValue.toString().trim())
//Log.Message(approvertable.getItem(0).getText_2(8).OleValue.toString().trim())
  if(approvertable.getItem(0).getText_2(6).OleValue.toString().trim()=="Rejected"){
  ValidationUtils.verify(true,true,"Purchase Order is Rejected by :"+Project_manager)
  TextUtils.writeLog("Purchase Order is Rejected by :"+Project_manager); 
  }else{ 
  TextUtils.writeLog("Purchase Order is Rejected by :"+Project_manager+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Purchase Order is Rejected by :"+Project_manager+ "But its Not Reflected")
  }
  ReportUtils.logStep_Screenshot();
  WorkspaceUtils.waitForObj(approvertable);
  ImageRepository.ImageSet.Undo.Click();
  aqUtils.Delay(1000);
  approvertable.Keys("[Tab]");
  aqUtils.Delay(1000);
  remarks.setText(" ");
  save.HoverMouse();
  ReportUtils.logStep_Screenshot();
  save.Click();
aqUtils.Delay(1000, "Waiting to Save");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
}


/*
var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.Action;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
  
  Delay(3000);
  Sys.Process("Maconomy").Refresh();
  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
  Sys.HighlightObject(table);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  ReportUtils.logStep_Screenshot();
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  Delay(4000);
  */


//}


var listPurchaseOrder = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.listPurchaseOrder.TabControl;
listPurchaseOrder.Click();
//aqUtils.Delay(3000, Indicator.Text);;
ImageRepository.ImageSet.Forward.Click();
//aqUtils.Delay(4000, Indicator.Text);;
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
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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
//WorkspaceUtils.closeAllWorkspaces();
}


//function CredentialLogin(){ 
//for(var i=level;i<Approve_Level.length;i++){
//  var UserN = true;
//  var temp="";
//  var Cred = Approve_Level[i].split("*");
//  for(var j=2;j<4;j++){
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
//  { 
//
//     var sheetName = "Agency Users";
//     workBook = Project.Path+excelName;
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//  }
//  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
//  { 
//
//    var sheetName = "SSC Users";
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
//  }
//  else{ 
//   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
//    if(UserN){ 
//      goToHR();
//      UserN = false;
//    }
//    temp = searchNumber(Eno);
//  }
////  Log.Message(temp)
//  if(temp.length!=0){
//    temp = temp+"*"+j;
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
//  break;
//  }
//  }
//  if((temp=="")||(temp==null))
//  Log.Error("User Name is Not available for level :"+i);
////  Log.Message("Logins :"+temp);
//}
//WorkspaceUtils.closeAllWorkspaces();
//
//}


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
if((temp.indexOf("Approve Purchase Order (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf("Approve Purchase Order (Substitute) (")!=-1)&&(temp1.length==3)){ 
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
if((temp.indexOf("Approve Purchase Order by Type (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf("Approve Purchase Order by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Purchase Order by Type (Substitute) from To-Dos List"); 
var listPass = false;   
  }
} 
  }
  
}


//function todo(lvl){ 
//  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
//  var toDo = Aliases.Maconomy.Shell.Composite.Composite.Composite.TodoGrid.PTabFolder.TabFolderPanel.ToDo;
//  toDo.HoverMouse();
//  ReportUtils.logStep_Screenshot();
//  toDo.DBlClick();
//  TextUtils.writeLog("Entering into To-Dos List");
//  aqUtils.Delay(3000, Indicator.Text);
//  //To Maximaize the window
//  Sys.Desktop.KeyDown(0x12);
//  Sys.Desktop.KeyDown(0x20);
//  Sys.Desktop.KeyUp(0x12);
//  Sys.Desktop.KeyUp(0x20);
//  Sys.Desktop.KeyDown(0x58);
//  Sys.Desktop.KeyUp(0x58);  
//  aqUtils.Delay(1000, Indicator.Text);
//
//try{
//var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
//}
//catch(e){
//var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//}
//refresh.Click();
//aqUtils.Delay(15000, Indicator.Text);
//try{
//Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
//}
//catch(e){
//Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
//}
//
//aqUtils.Delay(5000, Indicator.Text);
//if(lvl==2)
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Purchase Order (")!=-1)&&(temp1.length==2)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);   
//TextUtils.writeLog("Entering into Approve Purchase Order from To-Dos List"); 
//  }
//}
//if(lvl==3)
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Purchase Order (Substitute) (")!=-1)&&(temp1.length==3)){ 
//Client_Managt.ClickItem("|"+temp);    
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);   
//TextUtils.writeLog("Entering into Approve Purchase Order (Substitute) from To-Dos List");  
//  }
//}  
////if(lvl==3){
////Client_Managt.ClickItem("|Approve Purchase Order (Substitute) (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Purchase Order (Substitute) (*)");
////}
////if(lvl==2){
////Client_Managt.ClickItem("|Approve Purchase Order (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Purchase Order (*)");
////}
//
//
//}


//var PONum = "1707109739";
//var Apvr = "1707 Finace";
//var lvl = "0";



function FinalApprovePO(PONum,Apvr,lvl){ 
//aqUtils.Delay(5000, Indicator.Text);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//aqUtils.Delay(2000, Indicator.Text);
//ImageRepository.ImageSet.Show_Filter.Click();
//}


var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid;
var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox;
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(PONum);
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter;
WorkspaceUtils.waitForObj(table);
aqUtils.Delay(3000, "Reading Data in table");;
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
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
//aqUtils.Delay(5000, Indicator.Text);;

var approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.Approvals;
//var approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl
WorkspaceUtils.waitForObj(approvalBar);
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
    aqUtils.Delay(3000, Indicator.Text);;
//    if(ImageRepository.ImageSet.Maximize.Exists()){
    ImageRepository.ImageSet.Maximize.Click();
//    }
//aqUtils.Delay(3000, Indicator.Text);;
//var POapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
var POapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl2
WorkspaceUtils.waitForObj(POapproval);
POapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
POapproval.Click();
//aqUtils.Delay(3000, Indicator.Text);;
//var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
Sys.HighlightObject(approvertable)
for(var h=0;h<lvl;h++){ 
  approvertable.Keys("[Down]");
}

//var remarks = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.PORemarks;
var remarks = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.SWTObject("McTextWidget", "")
WorkspaceUtils.waitForObj(remarks);
remarks.setText("Rejected");
//aqUtils.Delay(1000, Indicator.Text);;
//var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.RemarksSave;
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
WorkspaceUtils.waitForObj(save);
save.HoverMouse();
ReportUtils.logStep_Screenshot();
save.Click();
//aqUtils.Delay(3000, Indicator.Text);;
//ImageRepository.ImageSet.Rejected.HoverMouse();
ImageRepository.ImageSet.Rejected.Click();
ReportUtils.logStep_Screenshot();
aqUtils.Delay(2000, Indicator.Text);;


ValidationUtils.verify(true,true,"Purchase Order is Rejected by "+Apvr)
TextUtils.writeLog("Levels 0 has  Rejected the Purchase Order by "+Apvr);

var ii=0;
while ((approvertable.getItem(0).getText_2(6).OleValue.toString().trim()!="Rejected")&&(ii!=60))
{
  aqUtils.Delay(100);
  ii++;
  approvertable.Refresh();
}

  if(approvertable.getItem(0).getText_2(6).OleValue.toString().trim()=="Rejected"){
  ValidationUtils.verify(true,true,"Purchase Order is Rejected by :"+Apvr)
  TextUtils.writeLog("Purchase Order is Rejected by :"+Apvr); 
  }else{ 
  TextUtils.writeLog("Purchase Order is Rejected by :"+Apvr+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Purchase Order is Rejected by :"+Apvr+ "But its Not Reflected")
  }
  ReportUtils.logStep_Screenshot();
  WorkspaceUtils.waitForObj(approvertable);
  ImageRepository.ImageSet.Undo.Click();
  aqUtils.Delay(1000);
  approvertable.Keys("[Tab]");
  aqUtils.Delay(1000);
  remarks.setText(" ");
  save.HoverMouse();
  ReportUtils.logStep_Screenshot();
  save.Click();
aqUtils.Delay(1000, "Waiting to Save");; 







//var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//menuBar.HoverMouse();
//aqUtils.Delay(2000, Indicator.Text);;
//if(ImageRepository.ImageSet.Undo.Exists())
//ImageRepository.ImageSet.Undo.Click();
//ReportUtils.logStep_Screenshot();
//aqUtils.Delay(8000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//aqUtils.Delay(2000, Indicator.Text);;
//var remarks = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.PORemarks;
//remarks.setText(" ");
//aqUtils.Delay(1000, Indicator.Text);;
//var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.RemarksSave;
//save.HoverMouse();
//ReportUtils.logStep_Screenshot();
//save.Click();
  
/*
var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.Action;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
   Delay(3000);
  Sys.Process("Maconomy").Refresh();
  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
  Sys.HighlightObject(table);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  ReportUtils.logStep_Screenshot();
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  Delay(4000);
  */
  
//  ValidationUtils.verify(true,true,"Purchase Order is Rejected by "+Apvr)
//  TextUtils.writeLog("Purchase Order is Rejected by "+Apvr);
//  aqUtils.Delay(8000, Indicator.Text);;  
  
  
}


}