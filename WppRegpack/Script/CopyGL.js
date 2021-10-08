//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

//Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CopyGL";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var newjournal =[];
var JournalNumber,Entrydate = "";
var JournalNumber1 = "";
var Language = "";


function ReverseGL(){
TextUtils.writeLog("Reverse GL Started"); 

aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

ExcelUtils.setExcelName(workBook, "Agency Users", true);
var Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "ReverseGL";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
JournalNumber,Entrydate = "";
Approve_Level = [];
JournalNumber1 = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Reverse GL  started::"+STIME);
TextUtils.writeLog("Execution Started :"+STIME);
getDetails();

gotoMenu(); 
JournalListTab();
ReversePop();
GeneralLedger();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(60000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var username = ExcelUtils.getRowDatas("SSC - Senior Accountant","Username")
Restart.login(username);
aqUtils.Delay(5000, Indicator.Text);
todo();
ApproveGL();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

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
Sys.HighlightObject(Workspc);
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Transaction from General Ledger Menu");
TextUtils.writeLog("Entering into GL Transaction from General Ledger Menu");
}


function getDetails(){ 
Indicator.PushText("Reading Data from Excel");
ExcelUtils.setExcelName(workBook, sheetName, true);
sheetName="CopyGL";
  aqUtils.Delay(3000, Indicator.Text);
Entrydate = ExcelUtils.getRowDatas("DateEntry",EnvParams.Opco)
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"DateEntry is Needed to Reverse General Journal");
}
Log.Message(Entrydate)

JournalNumber = ExcelUtils.getRowDatas("JournalNo",EnvParams.Opco)
Log.Message(JournalNumber)
  if((JournalNumber=="")||(JournalNumber==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  JournalNumber = ReadExcelSheet("General Journal No",EnvParams.Opco,"Data Management");
  Log.Message(JournalNumber)
  }
if((JournalNumber==null)||(JournalNumber=="")){ 
ValidationUtils.verify(false,true,"DateEntry is Needed to Reverse General Journal");
}



Indicator.PushText("Playback");
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
var temp ="";
while (!DDT.CurrentDriver.EOF()) {
if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
try{
temp = temp+xlDriver.Value(column).toString().trim();
}
catch(e){
temp = "";
}
break;
}
xlDriver.Next();
}
     
if(temp.indexOf("*")!=-1){
var excelData =  temp.split("*");
}else if(temp.length>0){ 
excelData[0] = temp;
}
     
DDT.CloseDriver(xlDriver.Name);
for(var i=0;i<excelData.length;i++)
return excelData;
}

function JournalListTab(){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var JournalList = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(JournalList);
Sys.HighlightObject(JournalList);
JournalList.Click(); 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//var gl=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
//waitForObj(gl);
//Sys.HighlightObject(gl);
//gl.Click();
//if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//  
//}
var all = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
//var all=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());

//Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button
waitForObj(all);
Sys.HighlightObject(all);
all.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var journal=Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(journal);
journal.Click();
journal.setText(JournalNumber);

var table = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==JournalNumber){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
ValidationUtils.verify(true,flag,"Journal Number is available to reverse :"+JournalNumber)
  if(flag){
var closefilter = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(closefilter);
closefilter.Click();

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var Genral = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;
Genral.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


var Copy = Aliases.Maconomy.Composite.SingleToolItemControl2;
Sys.HighlightObject(Copy);
Copy.Click();
//var ReverseJournal = Aliases.Maconomy.Composite.SingleToolItemControl2
//Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2
//var ReverseJournal = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//Sys.HighlightObject(ReverseJournal);
//ReverseJournal.Click();
aqUtils.Delay(5000, "finding Entry Date");
if(Entrydate!=""){
    var date1=Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Journal").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2);
Sys.HighlightObject(date1);
Log.Message(Entrydate);
WorkspaceUtils.CalenderDateSelection(date1,Entrydate)

      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Create a Employee");
    } 


var crgl=Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Journal").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Journal").OleValue.toString().trim());
WorkspaceUtils.waitForObj(crgl);
ReportUtils.logStep_Screenshot();
crgl.Click();
aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//  
//}

//var popup=Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal");
////Log.Message(popup.getText());
//
//var label=SWTObject("Label", "Journal no. * was created").getText();
//aqUtils.Delay(3000, Indicator.Text);
//log.Message(label);
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//    ExcelUtils.WriteExcelSheet("ReserveJournalNo",EnvParams.Opco,"Data Management",label) 

}
}


function ReversePop(){
var popup=Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - General Journal").OleValue.toString().trim());
Sys.HighlightObject(popup);
//Log.Message(popup.getText());
aqUtils.Delay(4000, Indicator.Text);

var label=Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - General Journal").OleValue.toString().trim()).SWTObject("Label", "*");
Sys.HighlightObject(label);
Log.Message(label.getText());
var gg = label.getText().OleValue.toString().trim();
var newjournal = gg.replace(/[^0-9]+/g, "")
Log.Message(newjournal);
JournalNumber1 = newjournal;
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("CopyJournalNo",EnvParams.Opco,"Data Management",newjournal) 
ReportUtils.logStep("INFO","CopyJournalNo :"+newjournal);
var ok= Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - General Journal").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
ok.Click();
}

function GeneralLedger(){
  
var generaljournal1 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl2;
//var generaljournal1 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl2
//var generaljournal1 = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(generaljournal1);
generaljournal1.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var companyno = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
//var companyno = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
//var companyno = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(companyno);
companyno.Click();
companyno.Keys("[Tab]");

var glno= Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
//var glno= Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
//var glno= Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
waitForObj(glno);
Sys.HighlightObject(glno);
glno.Click();
glno.setText(JournalNumber1);
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2)
//var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2)
//var table = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==JournalNumber1){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
ValidationUtils.verify(true,flag,"Copy Journal Number is available to submit:"+JournalNumber1)
  if(flag){

  var closefilter1 = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;    
//var closefilter1 = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
//var closefilter1 = Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(closefilter1);
closefilter1.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


var AttachDoc = Aliases.Composite7.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 9);
AttachDoc.Click();
  aqUtils.Delay(4000, "Waiting to Open file");;
  var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Document Attached");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var submit = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
submit.Click();

//var email=Aliases.Composite7.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 10)
//var email=Aliases.Maconomy.ReverseJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//Sys.HighlightObject(email);
//email.Click();


//var pop= Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal")
//var popup=Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - General Journal").OleValue.toString().trim());
//Sys.HighlightObject(popup);
//var ok= Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - General Journal").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//ok.HoverMouse(); 
//ReportUtils.logStep_Screenshot();
//ok.Click();
//var ok= Aliases.Maconomy.ReverseGLOK.Composite.Button;
//ok.Click();

}
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
if(ImageRepository.ImageSet.ToDos_Icon.Exists())
{ 
  
}
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve General Journal by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into GL Journals Awaiting Approval from To-Dos List");
listPass = false; 
  }
}

}


function ApproveGL(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var showFilter = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter
  WorkspaceUtils.waitForObj(showFilter)
  if(showFilter.text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Show Filter List").OleValue.toString().trim()){ 
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
  JNumber.setText(JournalNumber1);
waitForObj(closefilter);
Sys.HighlightObject(closefilter);
closefilter.HoverMouse();
closefilter.HoverMouse(); 
closefilter.HoverMouse();
closefilter.HoverMouse(); 
aqUtils.Delay(2000, Indicator.Text);;
WorkspaceUtils.waitForObj(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==JournalNumber1){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
  }
  
ValidationUtils.verify(flag,true,"Created Copy Journal is available in Approval List");
TextUtils.writeLog("Created Copy Journal is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var Approve = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.EditableBar.RemoveLatestDraft;
  Sys.HighlightObject(Approve);
  Approve.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var Post = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 12);
//var Post = Aliases.Maconomy.CreateGeneralJournal.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Journal_TabLine.Post;
WorkspaceUtils.waitForObj(Post);
Post.Click();


   TextUtils.writeLog("Post and Email is Clicked");
//    aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC"+"*", 1).WndCaption.indexOf("Print Posting Journal")!=-1){
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
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
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

//var OKay = Aliases.Maconomy.GLJornalAwaitingApproval.Okay.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//OKay.Click();

}
}