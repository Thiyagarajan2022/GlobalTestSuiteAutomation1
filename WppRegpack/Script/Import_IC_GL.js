//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/**
 * This script to Import IC General Ledger
 * @author  : Muthu Kumar M
 * @version : 2.0
 * Modified Date(MM/DD/YYYY) :11/09/2021
*/


Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "InterCompany ImportBudgetModel";
var Project_manager,jornalNumber,STIME,transaction_No="";


//Main Function
function Import_IC_GL() {
  
TextUtils.writeLog("Import IC General Ledger Started"); 
Indicator.PushText("waiting for window to open");

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Import IC General Ledger script
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}



excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "InterCompany ImportBudgetModel";
STIME,jornalNumber,transaction_No = "";

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Import IC General Ledger started::"+STIME);


goTo_GL_Transaction();
import_Delimited_File();
submit_GL();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(60000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var username = ExcelUtils.getRowDatas("SSC - Intercompany","Username");

Restart.login(username);
aqUtils.Delay(5000, Indicator.Text);

goTo_GL_Transaction();
goTo_Invoice();

}




//1.Open the Workspace Client. 
//2.Go to: General Ledger > GL Transactions.
function goTo_GL_Transaction(){ 
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
  var General_Ledger;
for(var i=1;i<=childCC;i++){ 
General_Ledger = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(General_Ledger.isVisible()){ 
General_Ledger = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
General_Ledger.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
General_Ledger.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Transactions from General Ledger Menu");
TextUtils.writeLog("Entering into GL Transactions from General Ledger Menu");
}



//3.Click the "Import" action button
//4.Tick the "Internal Names" box and click "Import"
//5.Now select the import file you prepared.

function import_Delimited_File(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(3000,"Waiting for Maconomy to complete loading");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var import_Icon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
import_Icon.Click();

aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var internal_Names = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import General Journal").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
if(!internal_Names.getSelection())
internal_Names.Click();
aqUtils.Delay(4000, Indicator.Text);

var import_File = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import General Journal").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import").OleValue.toString().trim());
import_File.Click();

aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var sFolder = Project.Path+"TestResource\\Import Budget Model\\";
var sFileName = EnvParams.Opco+"_IC General Journal Template.txt";
//Finding File ia availble or NOT
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ }
else{
Log.Error("Could not create the folder " + sFolder);
}
}
  
aqUtils.Delay(4000, "Waiting to Open file");;
var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
WorkspaceUtils.waitForObj(dicratory);
dicratory.Keys(sFolder+sFileName);
aqUtils.Delay(3000, "Waiting to Open file");;
var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
Sys.HighlightObject(opendoc);
WorkspaceUtils.waitForObj(opendoc);
opendoc.HoverMouse();
ReportUtils.logStep_Screenshot();
    
opendoc.Click();
aqUtils.Delay(2000, "File is Impoerted");

var p = Sys.Process("Maconomy").Window("#32770", "Save file", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
 var Reportfile = Sys.Process("Maconomy").Window("#32770", "Save file", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
 Sys.HighlightObject(Reportfile);
 var FileName = sFolder+EnvParams.Opco+"_"+Reportfile.wText;
 Reportfile.Keys(FileName)
 aqUtils.Delay(2000, "Document is Saving");
saveAs.Click();
}
   
aqUtils.Delay(2000, "Waiting For Completion");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }    

aqUtils.Delay(2000, "Waiting For Completion");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
    
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - Import General Journal").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Transactions - Import General Journal").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();


}
       
aqUtils.Delay(2000, "Waiting For Completion");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
      


}


// Today Date Format
function TodayDate(){ 
    var date = new Date();
    date.setDate(date.getDate());
    var dd = date.getDate();
    var mm = date.getMonth()+1; 
    var yyyy = date.getFullYear();
    date = mm+'/'+dd+'/'+yyyy;
    Log.Message(date);
    return date;
}



//6.In the General Journal filter list, select the imported journal. 
//7.Review the journal lines and click the "Submit and Email" action button.

function submit_GL(){ 
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
var Company_No = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
Company_No.Click();
Company_No.setText(EnvParams.Opco);
aqUtils.Delay(4000, Indicator.Text);
Company_No.Keys("[Tab][Tab][Tab]");
aqUtils.Delay(4000, Indicator.Text);
var Submitted_By = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
Submitted_By.setText(Project_manager);
aqUtils.Delay(4000, Indicator.Text);
Submitted_By.Keys("[Tab]");
aqUtils.Delay(4000, Indicator.Text);
var Created_On = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
Created_On.Click();
Created_On.setText(TodayDate())
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if((table.getItem(v).getText_2(0).OleValue.toString().trim()==EnvParams.Opco)&&(table.getItem(v).getText_2(3).OleValue.toString().trim()==Project_manager)
  &&(table.getItem(v).getText_2(4).OleValue.toString().trim()==TodayDate())&&(table.getItem(v).getText_2(10).OleValue.toString().trim()!="")
  &&(table.getItem(v).getText_2(11).OleValue.toString().trim()!="")){ 
    flag=true;
    table.Keys("[Down]");
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
  }
  
  
var CloseFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
Sys.HighlightObject(CloseFilter);

if(flag){ 
CloseFilter.Click();  
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


  jornalNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
  Sys.HighlightObject(jornalNumber);
  jornalNumber.Click();
  jornalNumber = jornalNumber.getText().OleValue.toString().trim();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(2000, "Submitting General Ledger");
    
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
transaction_No = table.getItem(0).getText_2(26).OleValue.toString().trim();

var attach_Document = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 9)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
attach_Document.Click();

  aqUtils.Delay(4000, "Waiting to Open file");;
  var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Document Attached");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
submit.Click();
aqUtils.Delay(2000, "Journal is Submitting");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

}

}


//8.Login to Maconomy. Go to General Ledger > GL Transactions > Intercompany Invoicing
//9.Select the OpCo which is recharging the another OpCo.
//10.In the Intercompany Invoicing card pane, enter selection criteria to find the imported journal and tick the "Show lines" box. 
//11.Then, in the "Intercompany Entries" table pane, tick the "Marked for Invoicing" box on the entry you want to invoice. 
//12.Last, click on "Print Invoice".
function goTo_Invoice(){ 
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var intercompany_Invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
Sys.HighlightObject(intercompany_Invoice);
intercompany_Invoice.Click()

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Company_Num = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
Company_Num.Click();
Company_Num.Keys(EnvParams.Opco);


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);



var flag = false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()== EnvParams.Opco){ 
      flag=true;
      table.Keys("[Down]");
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  
  
if(flag){ 
  var CloseFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  Sys.HighlightObject(CloseFilter);
  CloseFilter.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var Show_Lines = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
Sys.HighlightObject(Show_Lines);
if(!Show_Lines.getSelection()){
Show_Lines.Click();
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var intercompanyEntries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");                           
var Intercompany_Entries = intercompanyEntries.FindAllChildren("text","Intercompany Entries",20000,true)
Log.Message(Intercompany_Entries.FullName);
//                           Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
var Intercompany_Entries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
Sys.HighlightObject(Intercompany_Entries);
Intercompany_Entries.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Waiting for maconomy to load");

var Entries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
Entries.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
aqUtils.Delay(5000,"Waiting for maconomy to load");
var flag = false;
  for(var v=0;v<Entries.getItemCount();v++){ 
    if(Entries.getItem(v).getText_2(4).OleValue.toString().trim()== transaction_No){ 
      flag=true;
      var CheckBox = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 4).SWTObject("Button", "");
      Sys.HighlightObject(CheckBox);
      if(!CheckBox.getSelection())
      CheckBox.Click();
    }
    else{ 
      var CheckBox = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 4).SWTObject("Button", "");
      Sys.HighlightObject(CheckBox);
      if(CheckBox.getSelection())
      CheckBox.Click();
      Entries.Keys("[Down]");
    }
  }
  
  if(flag){ 
    var CheckBox = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 4).SWTObject("Button", "");
    Sys.HighlightObject(CheckBox);
    CheckBox.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(3000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
    Sys.HighlightObject(Save);
    Save.Click();
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(3000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
    var print_InterCompany_Invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6);
    Sys.HighlightObject(print_InterCompany_Invoice);
    print_InterCompany_Invoice.Click();


    aqUtils.Delay(5000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    aqUtils.Delay(5000,"Waiting for maconomy to load");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      
  
  
    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
    aqUtils.Delay(5000, Indicator.Text);
    WorkspaceUtils.savePDF_localDirectory("PDF Import IC General Journal","Print Invoice");
    
//var SaveTitle = "";
//var sFolder = "";
//var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
//    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice"+"*", 1).WndCaption.indexOf("Invoice")!=-1){
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    
//    if(ImageRepository.PDF.ChooseFolder.Exists())
//    ImageRepository.PDF.ChooseFolder.Click();
//    else{ 
//      var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
//      WorkspaceUtils.waitForObj(window);
//      Sys.Desktop.KeyDown(0x12); //Alt
//      Sys.Desktop.KeyDown(0x73); //F4
//      Sys.Desktop.KeyUp(0x12); //Alt
//      Sys.Desktop.KeyUp(0x73); //F4
//    aqUtils.Delay(2000, Indicator.Text);
//    Sys.HighlightObject(pdf)
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x41); //A 
//    Sys.Desktop.KeyUp(0x12); 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x41);
//    }
//    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
//    aqUtils.Delay(2000, Indicator.Text);
//    SaveTitle = save.wText;
//    
//sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
//if (! aqFileSystem.Exists(sFolder)){
//if (aqFileSystem.CreateFolder(sFolder) == 0){ 
//    
//}
//else{
//Log.Error("Could not create the folder " + sFolder);
//}
//}
//save.Keys(sFolder+SaveTitle+".pdf");
////var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
////saveAs.Click();
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
//var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
//Sys.HighlightObject(p);
//var saveAs = p.FindChild("WndCaption", "&Save", 2000);
//if (saveAs.Exists)
//{ 
//saveAs.Click();
//}
//aqUtils.Delay(2000, Indicator.Text);
//if(ImageRepository.ImageSet.SaveAs.Exists()){
//var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
//conSaveAs.Click();
//}
//Sys.HighlightObject(pdf);
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//    }
//ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
//Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
//ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
//
//
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//ExcelUtils.WriteExcelSheet("PDF IC Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  

    
  }

}



}




