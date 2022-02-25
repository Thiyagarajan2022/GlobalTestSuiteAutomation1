﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT ObjectUtils
//USEUNIT ActionUtils

/** 
 * This script create Vendor Invoice
 * @author  : Muthu Kumar M
 * @version : 3.0
  * Modified Date(MM/DD/YYYY) : 02/21/2022
 */

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "ApproveVendorInvoice";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var InvoiceNo ="";
var vID_Status = true;
var Language = "";
var Maconomy_ParentAddress,Maconomy_Index = "";

// Main Function
function ApproveInvoice(){ 
  

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
aqUtils.Delay(1000, Indicator.Text);

//Checking Login to execute Approve Vendor Invoice
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ApproveVendorInvoice";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
InvoiceNo ="";
vID_Status = true;



STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
try{
getDetails();
goTo_Account_Payable(); 
invoiceAllocation();

var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);

for(var i=level;i<ApproveInfo.length;i++){
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

ActionUtils.ToDos_Selection(Maconomy_ParentAddress, level, temp[3], "Approve Invoice Allocation Line","Approve Invoice Allocation Line by Type","Approve Invoice Allocation Line by Type (Substitute)","Approve Invoice Allocation Line (Substitute)");

FinalApproveinvoice(temp[1],temp[2],i,temp[3]);
}
}catch(err){ 
  Log.Message(err);
}
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);

ExcelUtils.setExcelName(workBook, "Data Management", true);

InvoiceNo = ReadExcelSheet("Reverse CreditNote Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((InvoiceNo=="")||(InvoiceNo==null)){
  
InvoiceNo = ReadExcelSheet("CreditNote Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((InvoiceNo=="")||(InvoiceNo==null)){
  
InvoiceNo = ReadExcelSheet("Reverse Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((InvoiceNo=="")||(InvoiceNo==null)){
  
InvoiceNo = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((InvoiceNo=="")||(InvoiceNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
InvoiceNo = ExcelUtils.getRowDatas("Invoice NO",EnvParams.Opco)
}
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice NO is Needed to Approve Vendor Invoice");
}
else{ 
  ValidationUtils.verify(true,true,"Approving Vendor Invoice NO :"+InvoiceNo)
}

}
else{ 
  ValidationUtils.verify(true,true,"Approving Reverse Vendor Invoice NO :"+InvoiceNo)
}

}
else{ 
  ValidationUtils.verify(true,true,"Approving CreditNote Vendor Invoice NO :"+InvoiceNo)
}

}else{ 
  ValidationUtils.verify(true,true,"Approving Reverse CreditNote Vendor Invoice NO :"+InvoiceNo)
}

}

// Approving 2nd Created vendor invoice for Reverse, credit note, reverse credit note
function Approve_Vendor_Invoice(){ 
  

Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

TextUtils.writeLog("Approve Vendor Invoice Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ApproveVendorInvoice";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
InvoiceNo ="";
vID_Status = true;



STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 
try{
getDetails_Dependency();
goToJobMenuItem();
invoiceAllocation();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
vID_Status = true;
todo(temp[3],i,temp[1],temp[2]);
//FinalApproveinvoice(temp[1],temp[2],i,temp[3]);
}
}catch(err){ 
  Log.Message(err);
}
WorkspaceUtils.closeAllWorkspaces();
}


function getDetails_Dependency(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);

ExcelUtils.setExcelName(workBook, "Data Management", true);
InvoiceNo = ReadExcelSheet("Second Vendor Invoice NO",EnvParams.Opco,"Data Management");
if((InvoiceNo=="")||(InvoiceNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
InvoiceNo = ExcelUtils.getRowDatas("Invoice NO",EnvParams.Opco)
}
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice NO is Needed to Approve Vendor Invoice");
}
else{ 
  ValidationUtils.verify(true,true,"Approving Vendor Invoice NO :"+InvoiceNo)
}

}


function goTo_Account_Payable(){ 
  
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_AccountPayable_from_workspace(); //Select Account Payable Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());

ReportUtils.logStep("INFO", "Moved to AP Transactions from Accounts Payable Menu");
TextUtils.writeLog("Entering into AP Transactions from Accounts Payable Menu");
}




function invoiceAllocation(){ 
  
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


  var allocation = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Invoice Allocation");
  Sys.HighlightObject(allocation);
  allocation.Click();
  
  

var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
if(closefilter==null){
var showfilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Show Filter List");
  Sys.HighlightObject(showfilter);
  showfilter.Click();
}

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();


var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,2);
//var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
var companyNo = table.SWTObject("McValuePickerWidget", "");
Sys.HighlightObject(companyNo);
companyNo.Click();
aqUtils.Delay(1000, Indicator.Text);
companyNo.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
//var invoiceNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
var invoiceNumber = table.SWTObject("McTextWidget", "", 2);
invoiceNumber.Click();
invoiceNumber.setText(InvoiceNo);



ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();



//aqUtils.Delay(5000, Indicator.Text);
//var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,2);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
if(table.getItem(v).getText_2(6).OleValue.toString().trim()==InvoiceNo){ 
  table.Keys("[Down]");
flag=true;    
break;
}
else{ 
table.Keys("[Down]");
}
}
ValidationUtils.verify(flag,true,"Created Vendor Invoice is available in system");
TextUtils.writeLog("Created Vendor Invoice is available in system");

 if(flag){
  var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();
//  aqUtils.Delay(5000, Indicator.Text);

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//  var invoiceApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
  var invoiceApproval = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",4);
  invoiceApproval = invoiceApproval.SWTObject("TabControl", "");
  Sys.HighlightObject(invoiceApproval);
  invoiceApproval.Click();
  
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
  if(ImageRepository.ImageSet.Forward.Exists()){ 
   if(ImageRepository.ImageSet.Maximize.Exists()){
   ImageRepository.ImageSet.Maximize.Click();
   }
  }
  
  aqUtils.Delay(3000, Indicator.Text);;
  ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

  var purchaseApproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","All Approval Actions");
  Sys.HighlightObject(purchaseApproval);  
  purchaseApproval.Click();
  
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var ApproverTable = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
  var y=0;
  for(var ii=0;ii<3;ii++){
  for(var i=0;i<ApproverTable.getItemCount();i++){  
    
  var approvers="";
  if(ApproverTable.getItem(i).getText_2(8)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
    if(ii==ApproverTable.getItem(i).getText_2(1).OleValue.toString().trim()){
  approvers = EnvParams.Opco+"*"+InvoiceNo+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
  Log.Message("Approver level :" +i+ ": " +approvers);
  //      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
  Approve_Level[y] = approvers;
  y++;
  break;
      }
      
      }
}
}
TextUtils.writeLog("Finding approvers for Created Vendor Invoice");
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


//function CredentialLogin(){ 
//  var uniqui = [];
//  var u=0;
//for(var i=level;i<Approve_Level.length;i++){
//  var UserN = true;
//  var temp="";
//  var Cred = Approve_Level[i].split("*");
//  for(var j=2;j<4;j++){
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
//  { 
//
//     var sheetName = "Agency Users";
//     workBook = Project.Path+excelName;
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//  }
//  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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
//    if(u==0){
//    uniqui[u] = temp;
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
//    Log.Message(ApproveInfo[i])
//    u++;
//    }else{
//      var GM_Status = false;
//    for(var gm = 0;gm<uniqui.length;gm++){
//      if(uniqui[gm]!=temp){
//    uniqui[u] = temp;
//    u++;
//    temp = temp+"*"+j;
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
//    Log.Message(ApproveInfo[i])
//    GM_Status = true;
//    break;
//    }
//    }
//    if(GM_Status){
//    break;
//    }
//  }
//  }
//  }
//  if((temp=="")||(temp==null))
//  Log.Error("User Name is Not available for level :"+i);
//  Log.Message("Logins :"+temp);
//}
//WorkspaceUtils.closeAllWorkspaces();
//
//}



//function CredentialLogin(){ 
//for(var i=level;i<Approve_Level.length;i++){
//  var UserN = true;
//  var temp="";
//  var Cred = Approve_Level[i].split("*");
//  for(var j=2;j<4;j++){
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
//  { 
//
//     var sheetName = "Agency Users";
//     workBook = Project.Path+excelName;
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//  }
//  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
//  { 
//
//    var sheetName = "SSC Users";
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
//  }
////  else{ 
////   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
////    if(UserN){ 
////      goToHR();
////      UserN = false;
////    }
////    temp = searchNumber(Eno);
////  }
////  Log.Message(temp)
//  if(temp.length!=0){
//    temp = temp+"*"+j;
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
//  break;
//  }
//  }
//  if((temp=="")||(temp==null))
//  Log.Error("User Name is Not available for level :"+i);
//  Log.Message("Logins :"+temp);
//}
//WorkspaceUtils.closeAllWorkspaces();
//
//}


function todo(lvl,apLvl,comID,vID){ 
  TextUtils.writeLog("Entering into To-Dos List");
  var toDo = Aliases.Maconomy.Shell.Composite.Composite.Composite.TodoGrid.PTabFolder.TabFolderPanel.ToDo;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible)
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible)
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;

refresh.Click();
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
//aqUtils.Delay(15000, Indicator.Text);
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible)
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible)
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;

//if(EnvParams.Country.toUpperCase()=="INDIA")
//   Runner.CallMethod("IND_ApproveVendorInvoice.todo",lvl,apLvl);
//else{
var listPass = true;
if(lvl==3){
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Allocation Line").OleValue.toString().trim())!=-1)&&(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Substitute").OleValue.toString().trim())!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type (Substitute) from To-Dos List"); 
FinalApproveinvoice(comID,vID,apLvl,lvl);
if(!vID_Status){
listPass = false;
break;
}
  }
}  
//Client_Managt.ClickItem("|Approve Invoice Allocation Line by Type (Substitute) (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Invoice Allocation Line by Type (Substitute) (*)");
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type (Substitute) from To-Dos List");

}
if(lvl==2){
  
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Allocation Line").OleValue.toString().trim())!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type from To-Dos List"); 
FinalApproveinvoice(comID,vID,apLvl,lvl);
if(!vID_Status){
listPass = false;
break;
}
  }
} 

//Client_Managt.ClickItem("|Approve Invoice Allocation Line by Type (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Invoice Allocation Line by Type (*)");
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type from To-Dos List");
listPass = false;
}

if(listPass){
if(lvl==3){

for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Allocation Line").OleValue.toString().trim())!=-1)&&(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Substitute").OleValue.toString().trim())!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Invoice Allocation Line (Substitute) from To-Dos List"); 
FinalApproveinvoice(comID,vID,apLvl,lvl);
if(!vID_Status){
listPass = false;
break;
}
  }
}

//Client_Managt.ClickItem("|Approve Invoice Allocation Line (Substitute) (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Invoice Allocation Line (Substitute) (*)");
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line (Substitute) from To-Dos List");
}
if(lvl==2){
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Invoice Allocation Line").OleValue.toString().trim())!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Invoice Allocation Line from To-Dos List"); 
FinalApproveinvoice(comID,vID,apLvl,lvl);
if(!vID_Status){
listPass = false;
break;
}
  }
} 
//Client_Managt.ClickItem("|Approve Invoice Allocation Line (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Invoice Allocation Line (*)");
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line from To-Dos List");
}
}
//}


if(vID_Status)
{ 
 ValidationUtils.verify(false,true,"Created Vendor Invoice is available in Approval List"); 
}
}

function FinalApproveinvoice(InvoiceNo,Apvr,lvl){ 
  

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder;
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
if(closefilter==null)
{ 
var showfilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Show Filter List");
Sys.HighlightObject(showfilter);
showfilter.Click();
}



ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,3);
Sys.HighlightObject(table);
//var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McPopupPickerWidget;
var firstCell = table.SWTObject("McPopupPickerWidget", "");
Sys.HighlightObject(firstCell);
firstCell.Keys("[Tab][Tab][Tab][Tab]");
aqUtils.Delay(1000, Indicator.Text);;
var invoiceNumber = table.SWTObject("McTextWidget", "", 2);
invoiceNumber.Click();
invoiceNumber.setText(InvoiceNo);

var table = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,3);
//aqUtils.Delay(6000, Indicator.Text);;
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//aqUtils.Delay(6000, Indicator.Text);;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(5).OleValue.toString().trim()==InvoiceNo){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
if(flag){
vID_Status = false;
ValidationUtils.verify(flag,true,"Created Vendor Invoice is available in Approval List");
TextUtils.writeLog("Created Vendor Invoice is available in Approval List");
if(flag){ 
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
var closefilter = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");

closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
//aqUtils.Delay(5000, Indicator.Text);;
//var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl2;

//var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl4;

var st = false;

var Approve = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl",JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve all").OleValue.toString().trim())

Sys.HighlightObject(Approve)
Sys.HighlightObject(Approve);
Sys.HighlightObject(Approve)
if(Approve.text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve all").OleValue.toString().trim()){
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
//aqUtils.Delay(8000, "Approving Invoice");;

  
ValidationUtils.verify(true,true,"Vendor Invoice is Approved by "+Apvr)
//aqUtils.Delay(8000, Indicator.Text);;
TextUtils.writeLog("Vendor Invoice is Approved by "+Apvr);
if(Approve_Level.length==lvl+1){
var approvalBar = ActionUtils.getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",3,"Composite",1);
approvalBar = approvalBar.SWTObject("TabControl", "");
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
//    aqUtils.Delay(3000, Indicator.Text);;
    if(ImageRepository.ImageSet.Maximize.Exists()){
    ImageRepository.ImageSet.Maximize.Click();
    }
//aqUtils.Delay(3000, Indicator.Text);;

ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

//var invoiceapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
var invoiceapproval = ActionUtils.getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","All Approval Actions");
Sys.HighlightObject(invoiceapproval);
invoiceapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
invoiceapproval.Click();
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();

var approvertable = ActionUtils.getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2)
Sys.HighlightObject(approvertable);
ActionUtils.waitUntil_MaconomyScreen_loaded_Completely();
ReportUtils.logStep_Screenshot();
}

  ValidationUtils.verify(true,true,"Vendor Invoice is Approved by "+Apvr)

}
}
else{ 
 ValidationUtils.verify(true,false,"Approve all Button is not available in maconomy") ;
}
}

} 
else{ 
ReportUtils.logStep_Screenshot();
var menuBar = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

}