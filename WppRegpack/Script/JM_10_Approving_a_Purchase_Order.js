﻿//USEUNIT CreditNotePO
//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT ReverseCreditNote
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT ObjectUtils
//USEUNIT ActionUtils

/**
 * This script create PO for Job
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Modified Date :02/14/2021
*/


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ApprovePurchaseOrder";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var POnumber = "";
var Language = "";
var Project_manager = "";
var Maconomy_ParentAddress,Maconomy_Index = "";

//Main Function
function ApprovePurchaseOrder(){ 
TextUtils.writeLog("Approve Purchase Order Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute to reject purchase order
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);


var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ApprovePurchaseOrder";
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
//try{
getDetails();
goto_AccountPayable();
gettingApproval();
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
Workspace_Client.Click();
WorkspaceUtils.closeAllWorkspaces();
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);

for(var i=level;i<ApproveInfo.length;i++){
  level=i;
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
Workspace_Client.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);


Maconomy_Index = WorkspaceUtils.Maconomy_Parent;
Maconomy_Index++;
WorkspaceUtils.Maconomy_Parent = Maconomy_Index;
Log.Message(Maconomy_Index);

// Restarting maconomy with Approver Logins
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
var Macscreen = WorkspaceUtils.switch_Maconomy(temp[2])
if(Macscreen=="Screen Not Found"){
Restart.login(temp[2]);
}
aqUtils.Delay(5000, Indicator.Text);

Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(temp[2]);
//Refreshing To-Do's List to find Submitted Job
aqUtils.Delay(5000, Indicator.Text);

ActionUtils.ToDos_Selection(Maconomy_ParentAddress, level, temp[3], "Approve Purchase Order by Type", "Approve Purchase Order", "Approve Purchase Order by Type (Substitute)", "Approve Purchase Order (Substitute)")

FinalApprovePO(temp[1],temp[2],i);
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

//}
//  catch(err){
//    Log.Message(err);
//  }
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
  
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
// Getting PO Number to Approve
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

}




//Moving to Purchase Order
function goto_AccountPayable(){ 
  
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_AccountPayable_from_workspace(); //Select Account Payable Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());

ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable Menu");
TextUtils.writeLog("Entering into Purchase Orders from Accounts Payable Menu");
}



function gettingApproval(){ 

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
if(closefilter==null)
{ 
var showfilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Show Filter List");
Sys.HighlightObject(showfilter);
showfilter.Click();
}

var allPurchase = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"Button", "My POs");
Sys.HighlightObject(allPurchase);
allPurchase.Click();

////  var allPurchase = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
//  var allPurchase = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
//  Sys.HighlightObject(allPurchase);
//  allPurchase.Click();


//var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
  var companyNo = table.SWTObject("McValuePickerWidget", "");
//  var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.CompanyNo;
  Sys.HighlightObject(companyNo);
  companyNo.Click();
  companyNo.Keys("[Tab]");
  
  var purchaseNo = table.SWTObject("McTextWidget", "", 2);
//  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox
////  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid.PurchaseNumber;
  Sys.HighlightObject(purchaseNo);
  purchaseNo.Click();
  purchaseNo.setText(POnumber);
  
//  var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
//  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.PurchaseTable.McGrid;
  Sys.HighlightObject(table);
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
   var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
//  var closefilter =Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter
//  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite.SingleToolItemControl;
  Sys.HighlightObject(closefilter);
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
  var purchaseOrderApproval = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",3);
  purchaseOrderApproval = purchaseOrderApproval.SWTObject("TabControl", "");
//  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.Approvals
//  var purchaseOrderApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  Sys.HighlightObject(purchaseOrderApproval);
  purchaseOrderApproval.Click();

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  ImageRepository.ImageSet.Maximize.Click();

  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
  var purchaseApproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Purchase Order Approval");
//  var purchaseApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3
//  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab;
  Sys.HighlightObject(purchaseApproval);
  purchaseApproval.Click();
  aqUtils.Delay(2000, "Reading Data from table");;
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  
  var ApproverTable = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
//  var ApproverTable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(ApproverTable);
   var y=0;
       Project_manager = eval(Maconomy_ParentAddress).WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
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
}
}








function FinalApprovePO(PONum,Apvr,lvl){ 

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();



var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
if(closefilter==null)
{ 
var showfilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Show Filter List");
Sys.HighlightObject(showfilter);
showfilter.Click();
}



var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid", "2",3);
var firstCell = table.SWTObject("McTextWidget", "");

WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
firstCell.setText(PONum);
//WorkspaceUtils.waitForObj(table);
aqUtils.Delay(3000, "Reading Data from table");;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
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

if(flag)
{
ValidationUtils.verify(flag,true,"Created Purchase Order is available in Approval List");
TextUtils.writeLog("Created Purchase Order is available in Approval List");
}
else
{
 ValidationUtils.verify(flag,true,"Created Purchase Order is not available in Approval List");
TextUtils.writeLog("Created Purchase Order is not available in Approval List");
  
}
if(flag){ 
  
var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
//closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();



//Approve PO

var Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Approve");
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

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

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
  
//Approve Bar
/*
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
*/

var approvalBar = "";
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.isVisible()){
//approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl
//}else{
//approvalBar = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
//}

approvalBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",3);
  approvalBar = approvalBar.SWTObject("TabControl", "");
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
ImageRepository.ImageSet.Maximize.Click();

if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN"))
   Runner.CallMethod("IND_ApprovePurchaseOrder.ApprovalStatus");
else{

var POapproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","PO Approval");
//var POapproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
//var POapproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3
WorkspaceUtils.waitForObj(POapproval)
POapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
POapproval.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
var approvertable = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
//var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//var approvertable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
WorkspaceUtils.waitForObj(approvertable)
ReportUtils.logStep_Screenshot();

for(var i=0;i<approvertable.getItemCount();i++){   
var approvers="";
if(approvertable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(true,false,"Created PO is not Approved")
}
}

}
}
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Apvr)
}
}

} 