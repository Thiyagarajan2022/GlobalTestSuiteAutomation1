//USEUNIT EnvParams
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
 
function RejectInvoice(){ 
TextUtils.writeLog("Reject Vendor Invoice Started"); 
Indicator.PushText("waiting for window to reponse");
aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Junior AP","Username")
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
sheetName = "RejectVendorInvoice";
level =0;
Approve_Level = [];
ApproveInfo = [];
mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
STIME = "";
InvoiceNo ="";

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creating Vendor Invoice started::"+STIME);
getDetails();
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
todo(temp[3]);
FinalApproveinvoice(temp[1],temp[2],i,temp[3]);
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
Client_Managt.ClickItem("|AP Transactions");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|AP Transactions");
}

} 
aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to AP Transactions from Accounts Payable Menu");
TextUtils.writeLog("Entering into AP Transactions from Accounts Payable Menu");
}

function invoiceAllocation(){ 
  var allocation = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Budgeting;
  allocation.Click();
  aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Show_Filter.Exists()){
ImageRepository.ImageSet.Show_Filter.Click();
}
aqUtils.Delay(3000, Indicator.Text);

var companyNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
companyNo.Click();
aqUtils.Delay(1000, Indicator.Text);
companyNo.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
var invoiceNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
invoiceNumber.Click();
invoiceNumber.setText(InvoiceNo);

aqUtils.Delay(5000, Indicator.Text);
var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
if(table.getItem(v).getText_2(6).OleValue.toString().trim()==InvoiceNo){ 
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
  aqUtils.Delay(5000, Indicator.Text);
  var invoiceApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
  invoiceApproval.Click();
  if(ImageRepository.ImageSet.Forward.Exists()){ 
   if(ImageRepository.ImageSet.Maximize.Exists()){
   ImageRepository.ImageSet.Maximize.Click();
   }
  }
  
  aqUtils.Delay(3000, Indicator.Text);;
  var purchaseApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.TabControl
  purchaseApproval.Click();
  var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite5.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
  var approvers="";
  if(ApproverTable.getItem(i).getText_2(8)!="Approved"){
  approvers = EnvParams.Opco+"*"+InvoiceNo+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
  Log.Message("Approver level :" +i+ ": " +approvers);
  //      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
  Approve_Level[y] = approvers;
  y++;
      }
}

  }
}

function CredentialLogin(){ 
for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 

     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }
  else{ 
   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
    if(UserN){ 
      goToHR();
      UserN = false;
    }
    temp = searchNumber(Eno);
  }
//  Log.Message(temp)
  if(temp.length!=0){
    temp = temp+"*"+j;
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
  break;
  }
  }
  if((temp=="")||(temp==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message("Logins :"+temp);
}
WorkspaceUtils.closeAllWorkspaces();

}


function todo(lvl){ 
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

try{
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
catch(e){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
try{
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
catch(e){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}

if(EnvParams.Country.toUpperCase()=="INDIA")
   Runner.CallMethod("IND_ApprovePurchaseOrder.todo",lvl);
else{
if(lvl==3){
Client_Managt.ClickItem("|Approve Invoice Allocation Line by Type (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Invoice Allocation Line by Type (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Invoice Allocation Line by Type (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Invoice Allocation Line by Type (*)");
}
}

}

function FinalApproveinvoice(InvoiceNo,Apvr,lvl){ 
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McPopupPickerWidget;
firstCell.Keys("[Tab][Tab][Tab][Tab]");
aqUtils.Delay(2000, Indicator.Text);;
var invoiceNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
invoiceNumber.Click();
invoiceNumber.setText(InvoiceNo);
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  
aqUtils.Delay(6000, Indicator.Text);;
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

ValidationUtils.verify(flag,true,"Created Vendor Invoice is available in Approval List");
TextUtils.writeLog("Created Vendor Invoice is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
//var Reject = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl3;
var Reject = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl5
Sys.HighlightObject(Reject)
if(Reject.isEnabled()){ 
Reject.HoverMouse();
ReportUtils.logStep_Screenshot();
Reject.Click();
aqUtils.Delay(8000, Indicator.Text);;
var remarkStatus = Aliases.Maconomy.Shell7.Composite2.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
remarkStatus.Click();
remarkStatus.setText("Rejected");
var rejectButton = Aliases.Maconomy.Shell7.Composite2.Composite.Composite2.Composite.Button;
rejectButton.Click();
aqUtils.Delay(5000, Indicator.Text);;
ValidationUtils.verify(true,true,"Vendor Invoice is Rejected by "+Apvr)
TextUtils.writeLog("Vendor Invoice is Rejected by "+Apvr);
aqUtils.Delay(8000, Indicator.Text);;
if(Approve_Level.length==lvl+1){
var approvalBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl;
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
    aqUtils.Delay(3000, Indicator.Text);;
    if(ImageRepository.ImageSet.Maximize.Exists()){
    ImageRepository.ImageSet.Maximize.Click();
    }
aqUtils.Delay(3000, Indicator.Text);;


var invoiceapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
invoiceapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
invoiceapproval.Click();
aqUtils.Delay(3000, Indicator.Text);;
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
ReportUtils.logStep_Screenshot();
}

  ValidationUtils.verify(true,true,"Vendor Invoice is Rejected by "+Apvr)
  aqUtils.Delay(8000, Indicator.Text);;  
}
}
}