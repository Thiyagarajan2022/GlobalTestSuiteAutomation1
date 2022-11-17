﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "RejectVendorInvoice";
var level =0;
var Approve_Level = [];
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var InvoiceNo ="";
var vID_Status = true;
function RejectInvoice(){ 
TextUtils.writeLog("Reject Vendor Invoice Started"); 
Indicator.PushText("waiting for window to reponse");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

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
sheetName = "RejectVendorInvoice";
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
try{
getDetails();
goToJobMenuItem();
invoiceAllocation();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
for(var i=level;i<1;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
vID_Status = true;
todo(temp[3],i,temp[1],temp[2]);
//todo(temp[3]);
//FinalApproveinvoice(temp[1],temp[2],i,temp[3]);
break;
}
}catch(err){ 
  Log.Message(err);
}
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);

ExcelUtils.setExcelName(workBook, "Data Management", true);
InvoiceNo = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,"Data Management");

if((InvoiceNo=="")||(InvoiceNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
InvoiceNo = ExcelUtils.getRowDatas("Invoice NO",EnvParams.Opco)
}
if((InvoiceNo==null)||(InvoiceNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice NO is Needed to Approve Vendor Invoice");
}
}

function goToJobMenuItem(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AP Transactions").OleValue.toString().trim());
}

} 
//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to AP Transactions from Accounts Payable Menu");
TextUtils.writeLog("Entering into AP Transactions from Accounts Payable Menu");
}

function invoiceAllocation(){ 
    while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  var allocation = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
  WorkspaceUtils.waitForObj(allocation);
  allocation.Click();
  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
  WorkspaceUtils.waitForObj(closefilter);
//  aqUtils.Delay(3000, Indicator.Text);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//ImageRepository.ImageSet.Show_Filter.Click();
//}
if(closefilter.text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Show Filter List").OleValue.toString().trim()){
  closefilter.Click();
}

//aqUtils.Delay(3000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   Sys.Desktop.Keys("[Up]")
var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(companyNo);
companyNo.Click();
aqUtils.Delay(1000, Indicator.Text);
companyNo.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
var invoiceNumber = //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "")
invoiceNumber.Click();
invoiceNumber.setText(InvoiceNo);

var labels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McPagingWidget", "", 1);
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
    labels = labels.Child(i);
    break;
  }
}

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
WorkspaceUtils.waitForObj(labels);
var i=0;
while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
  Log.Message(labels.getText().OleValue.toString().trim())
 ValidationUtils.verify(true,true,"Maconomy is loading continously......") 
}

//aqUtils.Delay(5000, Indicator.Text);
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
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
  var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  var invoiceApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
  WorkspaceUtils.waitForObj(invoiceApproval);
  invoiceApproval.Click();
  if(ImageRepository.ImageSet.Forward.Exists()){ 
   if(ImageRepository.ImageSet.Maximize.Exists()){
   ImageRepository.ImageSet.Maximize.Click();
   }
  }
  
  aqUtils.Delay(3000, Indicator.Text);;
  var purchaseApproval = 
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.TabControl
  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.InitiatePO
  WorkspaceUtils.waitForObj(purchaseApproval);  
  purchaseApproval.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var ApproverTable = 
  //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.POTable.McGrid
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


//function todo(lvl){ 
//  var toDo = Aliases.Maconomy.Shell.Composite.Composite.Composite.TodoGrid.PTabFolder.TabFolderPanel.ToDo;
//  toDo.HoverMouse();
//  ReportUtils.logStep_Screenshot();
//  toDo.DBlClick();
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
//if(EnvParams.Country.toUpperCase()=="INDIA")
//   Runner.CallMethod("IND_ApprovePurchaseOrder.todo",lvl);
//else{
////if(lvl==3){
////Client_Managt.ClickItem("|Approve Invoice Allocation Line (Substitute) (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line (Substitute) (*)");
////}
////if(lvl==2){
////Client_Managt.ClickItem("|Approve Invoice Allocation Line (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line (*)");
////}
//
//
//var listPass = true;
//if(lvl==3){
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Invoice Allocation Line by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);  
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type (Substitute) from To-Dos List"); 
//  }
//}  
////Client_Managt.ClickItem("|Approve Invoice Allocation Line by Type (Substitute) (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line by Type (Substitute) (*)");
////TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type (Substitute) from To-Dos List");
//listPass = false;
//}
//if(lvl==2){
//  
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Invoice Allocation Line by Type (")!=-1)&&(temp1.length==2)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);  
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type (Substitute) from To-Dos List"); 
//  }
//} 
//
////Client_Managt.ClickItem("|Approve Invoice Allocation Line by Type (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line by Type (*)");
////TextUtils.writeLog("Entering into Approve Invoice Allocation Line by Type from To-Dos List");
//listPass = false;
//}
//if(listPass){
//if(lvl==3){
//
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Invoice Allocation Line (Substitute) (")!=-1)&&(temp1.length==3)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);  
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line (Substitute) from To-Dos List"); 
//  }
//}
//
////Client_Managt.ClickItem("|Approve Invoice Allocation Line (Substitute) (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line (Substitute) (*)");
////TextUtils.writeLog("Entering into Approve Invoice Allocation Line (Substitute) from To-Dos List");
//}
//if(lvl==2){
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Invoice Allocation Line (")!=-1)&&(temp1.length==2)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);  
//TextUtils.writeLog("Entering into Approve Invoice Allocation Line from To-Dos List"); 
//  }
//} 
////Client_Managt.ClickItem("|Approve Invoice Allocation Line (*)");
////ReportUtils.logStep_Screenshot(); 
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line (*)");
////TextUtils.writeLog("Entering into Approve Invoice Allocation Line from To-Dos List");
//}
//}
//
//}
//
//}

function FinalApproveinvoice(InvoiceNo,Apvr,lvl){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  
  }
  while(!ImageRepository.ImageSet.Tab_Icon.Exists()){  }
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
//ImageRepository.ImageSet.Show_Filter.Click();
var showFilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SWTObject("SingleToolItemControl", "", 2);
WorkspaceUtils.waitForObj(showFilter);
showFilter.Click();
}


//aqUtils.Delay(5000, Indicator.Text);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//aqUtils.Delay(2000, Indicator.Text);
//ImageRepository.ImageSet.Show_Filter.Click();
//}

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(table);
var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McPopupPickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.Keys("[Tab][Tab][Tab][Tab]");
aqUtils.Delay(1000, Indicator.Text);;
var invoiceNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
invoiceNumber.Click();
invoiceNumber.setText(InvoiceNo);
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  
var labels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McPagingWidget", "", 2);

WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
    labels = labels.Child(i);
    break;
  }
}
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
WorkspaceUtils.waitForObj(labels);
var i=0;
while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}


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
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;

var Reject = //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl5
Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.markBatchAccruals

//var st = false;
//  var Reject = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2;
//  Sys.HighlightObject(Reject)
//  for(var i=0;i<Reject.ChildCount;i++){ 
//    if((Reject.Child(i).isVisible())&&(Reject.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject all").OleValue.toString().trim())){
//      Reject = Reject.Child(i);
//      st = true;
//      break;
//    }
//  }
//  
//  if(!st) { 
// Reject = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2;
// Sys.HighlightObject(Reject)
// for(var i=0;i<Reject.ChildCount;i++){ 
//  if((Reject.Child(i).isVisible())&&(Reject.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject all").OleValue.toString().trim())){
//    Reject = Reject.Child(i);
//    break;
//  }
//}
//}
Sys.HighlightObject(Reject)
if(Reject.isEnabled()){ 
Reject.HoverMouse();
ReportUtils.logStep_Screenshot();
Reject.Click();
aqUtils.Delay(8000, Indicator.Text);;
var remarkStatus = Aliases.Maconomy.Shell7.Composite2.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
WorkspaceUtils.waitForObj(remarkStatus);
remarkStatus.Click();
remarkStatus.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject all").OleValue.toString().trim());
var rejectButton = Sys.Process("Maconomy").SWTObject("Shell", "Reject").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject").OleValue.toString().trim())
//Aliases.Maconomy.Shell7.Composite2.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reject all").OleValue.toString().trim())
rejectButton.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
aqUtils.Delay(2000, "Rejecting Invoice");; 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
//WorkspaceUtils.waitForObj(screen);
screen.Click();
screen.MouseWheel(-200);
var ApvPerson = "";
var Apv = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("Composite", "");
//var Apv = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 1);
for(var a=0;a<Apv.ChildCount;a++){ 
  if((Apv.Child(a).Visible)&&(Apv.Child(a).JavaClassName == "McTextWidget")){ 
    ApvPerson = Apv.Child(a);
    break;
  }
}
if((ApvPerson=="")||(ApvPerson==null)){ 
ApvPerson = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite3.SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);  
            
}


var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
ApvPerson.Click();
while ((((ApvPerson.getText().OleValue.toString().trim().indexOf("ejected")==-1)&&(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("YOU")==-1))&&(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))&&(i!=600))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

  Log.Message(ApvPerson.getText())
  Log.Message((ApvPerson.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "ejected").OleValue.toString().trim())!==-1))
  if(((ApvPerson.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "ejected").OleValue.toString().trim())!==-1))){
  ValidationUtils.verify(true,true,"Vendor Invoice is Rejected by :"+loginPer)
  TextUtils.writeLog("Vendor Invoice is Rejected by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Vendor Invoice is Rejected by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Vendor Invoice is Rejected by :"+loginPer+ "But its Not Reflected")
  }
  
  
  
  
  
  
//aqUtils.Delay(5000, Indicator.Text);;
ValidationUtils.verify(true,true,"Vendor Invoice is Rejected by "+Apvr)
TextUtils.writeLog("Vendor Invoice is Rejected by "+Apvr);
//aqUtils.Delay(8000, Indicator.Text);;
//if(Approve_Level.length==lvl+1){
var approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl;
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
//    aqUtils.Delay(3000, Indicator.Text);;
    if(ImageRepository.ImageSet.Maximize.Exists()){
    ImageRepository.ImageSet.Maximize.Click();
    }
//aqUtils.Delay(3000, Indicator.Text);;


var invoiceapproval = 
Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 12)  

WorkspaceUtils.waitForObj(invoiceapproval);
invoiceapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
invoiceapproval.Click();
//aqUtils.Delay(3000, Indicator.Text);;
var approvertable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.SWTObject("Composite", "", 11).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
//Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
WorkspaceUtils.waitForObj(approvertable);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
ReportUtils.logStep_Screenshot();
var closepanel = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.listPurchaseOrder.TabControl
closepanel.Click();
//aqUtils.Delay(3000, Indicator.Text);;
ImageRepository.ImageSet.Forward.Click();
//}

  ValidationUtils.verify(true,true,"Vendor Invoice is Rejected by "+Apvr)
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var undoReject = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.submitjournal
  
// var st = false; 
//  var undoReject = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2;
//  Sys.HighlightObject(undoReject)
//  for(var i=0;i<undoReject.ChildCount;i++){ 
//    if((undoReject.Child(i).isVisible())&&(undoReject.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Undo All Approvals/Rejections").OleValue.toString().trim())){
//      undoReject = undoReject.Child(i);
//      st = true;
//      break;
//    }
//  }
//  
//  if(!st) { 
// undoReject = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2;
// Sys.HighlightObject(undoReject)
//  for(var i=0;i<undoReject.ChildCount;i++){ 
//    if((undoReject.Child(i).isVisible())&&(undoReject.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Undo All Approvals/Rejections").OleValue.toString().trim())){
//      undoReject = undoReject.Child(i);
//      break;
//    }
//  }
//  }
  undoReject.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
}
}

} 
else{ 
ReportUtils.logStep_Screenshot();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}
}