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
var sheetName = "CreateGeneralJournal";
var GRP_1,VC1,Job_Number1,work1,Debit_1,Credit_1,GRP_2,VC2,Job_Number2,work2,Debit_2,Credit_2,Job_Type,department,buss_unit = ""
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var jornalNumber = "";

//Main Function
function CreateGeneralJournal(){ 
TextUtils.writeLog("Create General Journal Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
Project_manager = ExcelUtils.getRowDatas("SSC - Junior Accountant","Username")
Log.Message(Project_manager);
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
sheetName = "CreateGeneralJournal";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
GRP_1,VC1,Job_Number1,work1,Debit_1,Credit_1,GRP_2,VC2,Job_Number2,work2,Debit_2,Credit_2,Job_Type,department,buss_unit = ""
jornalNumber = "";
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create General Journal started::"+STIME);

getDetails();
gotoMenu();
gotoGeneralJournal();
AddJournalLines();
attachDocument();
submit();
WorkspaceUtils.closeAllWorkspaces();
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var username = ExcelUtils.getRowDatas("SSC - Senior Accountant","Username")
Restart.login(username);
aqUtils.Delay(5000, Indicator.Text);
todo();
ApproveGL();
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 

ExcelUtils.setExcelName(workBook, sheetName, true);
GRP_1 = ExcelUtils.getColumnDatas("GRP_1",EnvParams.Opco)

if((GRP_1==null)||(GRP_1=="")){ 
ValidationUtils.verify(false,true,"GRP_1 is Needed to Create General Journal");
}


if(GRP_1=="P"){
ExcelUtils.setExcelName(workBook, "Data Management", true);
VC1 = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
if((VC1=="")||(VC1==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VC1 = ExcelUtils.getColumnDatas("Vendor/Client_1",EnvParams.Opco)
}
if((VC1==null)||(VC1=="")){ 
ValidationUtils.verify(false,true,"Vendor Number(Vendor/Client_1) is Needed to Create General Journal");
}
}

if(GRP_1=="R"){
ExcelUtils.setExcelName(workBook, "Data Management", true);
VC1 = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
if((VC1=="")||(VC1==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VC1 = ExcelUtils.getColumnDatas("Vendor/Client_1",EnvParams.Opco)
}
if((VC1==null)||(VC1=="")){ 
ValidationUtils.verify(false,true,"Client Number(Vendor/Client_1) is Needed to Create General Journal");
}
}

if((GRP_1=="P")||(GRP_1=="R")){
ExcelUtils.setExcelName(workBook, "Data Management", true);
Job_Number1 = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
}else{ 

if((Job_Number1=="")||(Job_Number1==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_Number1 = ExcelUtils.getColumnDatas("Job Number_1",EnvParams.Opco)
}
if((Job_Number1==null)||(Job_Number1=="")){ 
ValidationUtils.verify(false,true,"Job Number 1 is Needed to Create General Journal");
}

}

if((GRP_1=="P")||(GRP_1=="R")){
for(var i=1;i<=10;i++){
sheetName = "JobBudgetCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
 work1 = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
 Log.Message(work1);
if((work1!="")&&(work1!=null)&&(work1.indexOf("T")==-1)){
 break;
}else{ 
 work1 = ""; 
}
}
}else{ 

ExcelUtils.setExcelName(workBook, sheetName, true);
if((work1==null)||(work1=="")){ 
work1 = ExcelUtils.getColumnDatas("WorkCode_1",EnvParams.Opco)
}
if((work1==null)||(work1=="")){ 
ValidationUtils.verify(false,true,"Workcode_1 is Needed to Create General Journal");
}

}

ExcelUtils.setExcelName(workBook, sheetName, true);
Debit_1 = ExcelUtils.getColumnDatas("Debit_1",EnvParams.Opco)
Credit_1 = ExcelUtils.getColumnDatas("Credit_1",EnvParams.Opco)
if(((Debit_1==null)||(Debit_1==""))&&((Credit_1==null)||(Credit_1==""))){ 
ValidationUtils.verify(false,true,"Debit_1 or Credit_1 is Needed to Create General Journal");
}

//Line 2
ExcelUtils.setExcelName(workBook, sheetName, true);
GRP_2 = ExcelUtils.getColumnDatas("GRP_2",EnvParams.Opco)

if((GRP_2==null)||(GRP_2=="")){ 
ValidationUtils.verify(false,true,"GRP_2 is Needed to Create General Journal");
}


if(GRP_2=="P"){
ExcelUtils.setExcelName(workBook, "Data Management", true);
VC2 = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
if((VC2=="")||(VC2==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VC2 = ExcelUtils.getColumnDatas("Vendor/Client_2",EnvParams.Opco)
}
if((VC2==null)||(VC2=="")){ 
ValidationUtils.verify(false,true,"Vendor Number(Vendor/Client_2) is Needed to Create General Journal");
}
}

if(GRP_2=="R"){
ExcelUtils.setExcelName(workBook, "Data Management", true);
VC2 = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
if((VC2=="")||(VC2==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VC2 = ExcelUtils.getColumnDatas("Vendor/Client_2",EnvParams.Opco)
}
if((VC2==null)||(VC2=="")){ 
ValidationUtils.verify(false,true,"Client Number(Vendor/Client_2) is Needed to Create General Journal");
}
}

if((GRP_2=="P")||(GRP_2=="R")){
ExcelUtils.setExcelName(workBook, "Data Management", true);
Job_Number2 = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
}else{ 

if((Job_Number2=="")||(Job_Number2==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_Number2 = ExcelUtils.getColumnDatas("Job Number_2",EnvParams.Opco)
}
if((Job_Number2==null)||(Job_Number2=="")){ 
ValidationUtils.verify(false,true,"Job Number 2 is Needed to Create General Journal");
}

}

if((GRP_2=="P")||(GRP_2=="R")){
for(var i=1;i<=10;i++){
sheetName = "JobBudgetCreation";
ExcelUtils.setExcelName(workBook, sheetName, true);
 work2 = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
 Log.Message(work2);
if((work2!="")&&(work2!=null)&&(work2.indexOf("T")==-1)){
 break;
}else{ 
 work2 = ""; 
}
}
}else{ 

ExcelUtils.setExcelName(workBook, sheetName, true);
if((work2==null)||(work2=="")){ 
work2 = ExcelUtils.getColumnDatas("WorkCode_2",EnvParams.Opco)
}
if((work2==null)||(work2=="")){ 
ValidationUtils.verify(false,true,"Workcode_2 is Needed to Create General Journal");
}

}

ExcelUtils.setExcelName(workBook, sheetName, true);
Debit_2 = ExcelUtils.getColumnDatas("Debit_2",EnvParams.Opco)
Credit_2 = ExcelUtils.getColumnDatas("Credit_2",EnvParams.Opco)
if(((Debit_2==null)||(Debit_2==""))&&((Credit_2==null)||(Credit_2==""))){ 
ValidationUtils.verify(false,true,"Debit_2 or Credit_2 is Needed to Create General Journal");
}

ExcelUtils.setExcelName(workBook, "JobCreation", true);
Job_Type = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
if((Job_Type==null)||(Job_Type=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Job_Type = ExcelUtils.getColumnDatas("Job Type",EnvParams.Opco)
}
if((Job_Type==null)||(Job_Type=="")){ 
ValidationUtils.verify(false,true,"Job Type Number is Needed to Create General Journal");
}

ExcelUtils.setExcelName(workBook, "JobCreation", true);
department = ExcelUtils.getRowDatas("Department",EnvParams.Opco)
if((department==null)||(department=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
department = ExcelUtils.getColumnDatas("Job Department",EnvParams.Opco)
}
if((department==null)||(department=="")){ 
ValidationUtils.verify(false,true,"Department Number is Needed to Create General Journal");
}

ExcelUtils.setExcelName(workBook, "JobCreation", true);
buss_unit = ExcelUtils.getRowDatas("BusinessUnit",EnvParams.Opco)
if((buss_unit==null)||(buss_unit=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
buss_unit = ExcelUtils.getColumnDatas("Job BusinessUnit",EnvParams.Opco)
}
if((buss_unit==null)||(buss_unit=="")){ 
ValidationUtils.verify(false,true,"BusinessUnit Number is Needed to Create General Journal");
}

Log.Message(Job_Number1)
Log.Message(Job_Number2)
 
}


function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.GendralLedger.Exists()){
ImageRepository.ImageSet.GendralLedger.Click();// GL
}
else if(ImageRepository.ImageSet.GendralLedger1.Exists()){
ImageRepository.ImageSet.GendralLedger1.Click();
}
else{
ImageRepository.ImageSet.GendralLedger2.Click();
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
Client_Managt.ClickItem("|GL Transactions");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|GL Transactions");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Transactions from General Ledger Menu");
TextUtils.writeLog("Entering into GL Transactions from General Ledger Menu");
}

function gotoGeneralJournal(){ 
  
var table = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Visible){
var closeFilter = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter;
WorkspaceUtils.waitForObj(closeFilter);
closeFilter.Click();
}

var Tabfolder = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine;
WorkspaceUtils.waitForObj(Tabfolder);
for(var i=0;i<Tabfolder.ChildCount;i++){ 
  if((Tabfolder.Child(i).isVisible())&&(Tabfolder.Child(i).toolTipText=="New Journal (Ctrl+N)")){
    Tabfolder.Child(i).Click();
    break;
  }
}

//var NewJournal = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine.NewJournal;

var company = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.CompanyNumber;
WorkspaceUtils.waitForObj(company);
  if(EnvParams.Opco!=""){
  company.Click();
  WorkspaceUtils.SearchByValue(company,"Company",EnvParams.Opco,"Company");
    }else{ 
    ValidationUtils.verify(false,true,"Company Number is Need to create PurchaseOrder");
  }
  
var Save = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine.Save;
WorkspaceUtils.waitForObj(Save);
Save.Click();

}

function AddJournalLines() {
  for(var i=0;i<2;i++){
var addline = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.AddLine;
WorkspaceUtils.waitForObj(addline);
var ii=0;
while ((addline.toolTipText!="Add Entry (Ctrl+M)")&&(ii<60))
{
  aqUtils.Delay(100);
  ii++;
  addline.Refresh();
}
addline.Click();

var entryDate = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.EntryDate;
WorkspaceUtils.waitForObj(entryDate);
entryDate.Keys("[Tab][Tab]");

var GRP = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.GRP;
WorkspaceUtils.waitForObj(GRP);
GRP.Keys(GRP_1);
var ii=0;
while ((GRP.getText().OleValue.toString().trim()!=GRP_1)&&(ii<60))
{
  aqUtils.Delay(100);
  ii++;
  GRP.Refresh();
}
GRP.Keys("[Tab]");
var VC = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.VendorCustomer_No;
if(GRP_1=="R"){ 
VC.HoverMouse();
Sys.HighlightObject(VC); 
if(VC1!=""){
VC.Click();
WorkspaceUtils.VPWSearchByValue(VC,"Client",VC1,"Client Number");
}
}

if(GRP_1=="P"){ 
VC.HoverMouse();
Sys.HighlightObject(VC); 
if(VC1!=""){
VC.Click();
  SearchByValues_Col_1_all(vendor,"Vendor",VendorID,"Vendor Number","All Vendors");
  }
}

Sys.HighlightObject(VC); 
VC.Keys("[Tab][Tab]");
var Job = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.VendorCustomer_No;
Job.Click();

if(i==0)
WorkspaceUtils.SearchByValues_all_Col_2(Job,"Job",Job_Number1,"Job Number","All Jobs");
if(i==1)
WorkspaceUtils.SearchByValues_all_Col_2(Job,"Job",Job_Number2,"Job Number","All Jobs");
Job.Keys("[Tab]");

var work = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.VendorCustomer_No;
work.Click();
if(i==0)
WorkspaceUtils.SearchByValue(work,"Work Code",work1,"WorkCode");
if(i==1)
WorkspaceUtils.SearchByValue(work,"Work Code",work2,"WorkCode");
work.Keys("[Tab][Tab]");

var Debit = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.Amount;
Debit.Click();
if(i==0){ 
  if((Debit_1!="")&&(Debit_1!=null))
  Debit.setText(Debit_1)
}

if(i==1){ 
  if((Debit_2!="")&&(Debit_2!=null))
  Debit.setText(Debit_2)
}
Debit.Keys("[Tab]");
var credit = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.Amount;
credit.Click();
if(i==0){ 
  if((Credit_1!="")&&(Credit_1!=null))
  credit.setText(Credit_1)
}

if(i==1){ 
  if((Credit_2!="")&&(Credit_2!=null))
  credit.setText(Credit_2)
}
credit.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");

var JType = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.VendorCustomer_No;
WorkspaceUtils.SearchByValue(JType,"Local Specification 1",Job_Type,"Job Type");
JType.Keys("[Tab]");
var JDepartment = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.VendorCustomer_No;
WorkspaceUtils.SearchByValue(JDepartment,"Local Specification 2",department,"Job Department");
JDepartment.Keys("[Tab][Tab]");
var JBusinessUnit = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.VendorCustomer_No;
WorkspaceUtils.SearchByValue(JBusinessUnit,"Local Specification 4",buss_unit,"Job BusinessUnit");

var save = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.Save;
WorkspaceUtils.waitForObj(save);
save.Click();
}
}


function attachDocument(){ 

  var doc =  Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine.AttachDocument;

  Sys.HighlightObject(doc);
  doc.HoverMouse();
  doc.HoverMouse();
  doc.HoverMouse();
  doc.Click();
//  var attchDocument = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.Composite.AttachDocument;
//  WorkspaceUtils.waitForObj(attchDocument)
//  ReportUtils.logStep_Screenshot();
//  attchDocument.Click();
  aqUtils.Delay(4000, "Waiting to Open file");;
  var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Document Attached");
}


function submit(){ 
  jornalNumber = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.JournalNumber;
  Sys.HighlightObject(jornalNumber);
  jornalNumber = jornalNumber.getText().OleValue.toString().trim();
  
  var submit = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine.Submit;
  WorkspaceUtils.waitForObj(submit);
  submit.Click();
  
  var Okay = Aliases.Maconomy.GLJornalAwaitingApproval.Okay.Button;
  Okay.Click();
  

  ValidationUtils.verify(true,true,"Journal Number :"+jornalNumber);
 
}

function todo(){ 
  TextUtils.writeLog("Loged into SSC - Senior Accountant Approver login"); 
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

for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("GL Journals Awaiting Approval (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into GL Journals Awaiting Approval from To-Dos List");
listPass = false; 
  }
}

}


function ApproveGL(){ 
  var showFilter = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter
  WorkspaceUtils.waitForObj(showFilter)
  if(showFilter.text=="Show Filter List"){ 
    showFilter.Click();
  }
  var table = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget2.McGrid;
  WorkspaceUtils.waitForObj(table);
  var firstcell = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget2.McGrid.McValuePickerWidget;
  var closefilter = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.setText(EnvParams.Opco);
  firstcell.Keys("[Tab]");
  var JNumber = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget2.McGrid.JournalNumber;
  WorkspaceUtils.waitForObj(JNumber);
  JNumber.Click();
  JNumber.setText(jornalNumber);
waitForObj(closefilter);
Sys.HighlightObject(closefilter);
closefilter.HoverMouse();
closefilter.HoverMouse(); 
closefilter.HoverMouse();
closefilter.HoverMouse(); 
aqUtils.Delay(2000, Indicator.Text);;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==jornalNumber){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
  }
  
ValidationUtils.verify(flag,true,"Created General Journal is available in Approval List");
TextUtils.writeLog("Created General Journal is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

var Post = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine.Post;
WorkspaceUtils.waitForObj(Post);
Post.Click();

   TextUtils.writeLog("Post and Email is Clicked");
//    aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Posting Journal")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x41);
    
    if(ImageRepository.PDF.ChooseFolder.Exists())
    ImageRepository.PDF.ChooseFolder.Click();
    
    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
    aqUtils.Delay(2000, Indicator.Text);
    SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
saveAs.Click();
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Posting Journal is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);

var OKay = Aliases.Maconomy.GLJornalAwaitingApproval.Okay.Button;
OKay.Click();

  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("General Journal No",EnvParams.Opco,"Data Management",jornalNumber);
  TextUtils.writeLog("General Journal No :"+jornalNumber); 
  
}
}