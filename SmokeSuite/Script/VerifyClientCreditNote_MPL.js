//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

/** 
 * This script implements printing of Client Credit Note PDF and verify contents against Maconomy
 * @author  : Sai Kiran Vemula
 * @version : 1.0
 * Created Date :08/11/2020
 */
var Project_manager = "";
var Language = "";

/**
  *  This function invokes maconomy and calls subfunctionality methods
  */
  
function verifyClientCreditNoteMPL(){
TextUtils.writeLog("Verification Of Client Credit Note MPL"); 

//Setting Language in WorkspaceUtils
Language = "";
Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

var excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "Data Management";
ExcelUtils.setExcelName(workBook, sheetName, true);
jobNumber = ExcelUtils.getRowDatas("Client_CreditNote_JobNumber",EnvParams.Opco)

//Checking Login for Client Creation
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
}

try{
WorkspaceUtils.closeAllWorkspaces();
gotoMenu(); 
verify_ClientCreditNoteMPL(jobNumber);
}
  catch(err){
    Log.Message(err);
  }
}

/**
  *  This function navigates to Jobs under Jobs Workspace
  */

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();

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
var scroll = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "")
scroll.Click();
scroll.MouseWheel(500);
if(ImageRepository.ImageSet.JobsMenu.Exists())
ImageRepository.ImageSet.JobsMenu.Click();
else if (ImageRepository.ImageSet.Jobs_Workspace.Exists())
ImageRepository.ImageSet.Jobs_Workspace.Click();
else{
     ReportUtils.logStep("Fail", "Jobs Workspace not displayed");
     Log.Message("Jobs Workspace not displayed");
}

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|Jobs");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Jobs");
}
}

if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Worspace");
Log.Message("Entering Jobs from Jobs Workspace");
}
}

/**
  *  This function prints Client Credit Note PDF and validates fields
  */

function verify_ClientCreditNoteMPL(jobNumber){ 

var jobTab= Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(jobTab);
Sys.HighlightObject(jobTab);

var allJobsRadio = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
Sys.HighlightObject(allJobsRadio);
allJobsRadio.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var table = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

var firstCell = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.Keys("[Tab][Tab]");
aqUtils.Delay(1000, Indicator.Text);;
var jobNumberField = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
jobNumberField.Click();
jobNumberField.setText(jobNumber);
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var closeFilter = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
closeFilter.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var InvoicingTab = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(InvoicingTab);
Sys.HighlightObject(InvoicingTab);
InvoicingTab.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var draftInvoicesTab = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
WorkspaceUtils.waitForObj(draftInvoicesTab);
Sys.HighlightObject(draftInvoicesTab);
draftInvoicesTab.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var closeSecondFilter = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closeSecondFilter.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists()){}

var clientName = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
clientName = clientName.getText().OleValue.toString().trim();


var printCreditMemoBtn = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.SingleToolItemControl2
WorkspaceUtils.waitForObj(printCreditMemoBtn);  
Sys.HighlightObject(printCreditMemoBtn);
ReportUtils.logStep_Screenshot("");
printCreditMemoBtn.Click();

//var printDraftBtn = Aliases.Maconomy.Jobs_Invoice_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.SingleToolItemControl;
//WorkspaceUtils.waitForObj(printDraftBtn);  
//printDraftBtn.Click();
TextUtils.writeLog("Print Client Credit Note is Clicked and saved"); 
if(ImageRepository.ImageSet.LoadedBox.Exists()){}
aqUtils.Delay(5000, Indicator.Text);


var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Client Credit Note")!=-1){
   aqUtils.Delay(2000, Indicator.Text);
WorkspaceUtils.waitForObj(pdf);
Sys.HighlightObject(pdf)
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
    
if(ImageRepository.PDF.ChooseDiffFolder.Exists())
ImageRepository.PDF.ChooseDiffFolder.Click();
else{ 
var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
WorkspaceUtils.waitForObj(window);

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x73); //F4
Sys.Desktop.KeyUp(0x12); //Alt
Sys.Desktop.KeyUp(0x73); //F4
aqUtils.Delay(2000, Indicator.Text);
Sys.HighlightObject(pdf)

Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
}

aqUtils.Delay(2000, "Waiting for SaveAs Window");
var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
aqUtils.Delay(3000, Indicator.Text);
SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
var completePdfPath = sFolder+SaveTitle+".pdf";
save.Keys(sFolder+SaveTitle+".pdf");
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
saveAs.Click();
aqUtils.Delay(4000, Indicator.Text);

Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print Draft is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");


  var docObj;
  // Load the PDF file to the PDDocument object
  try{
  Log.Message(completePdfPath)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(completePdfPath);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
  var pdflineSplit = docObj.split("\r\n");
 
  for (j=0; j<pdflineSplit.length; j++)
  {
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Note No").OleValue.toString().trim()))
    {   
        Log.Message("Credit Note Number is available in Pdf")
        ReportUtils.logStep_Screenshot("");
        ValidationUtils.verify(true,true,"Credit Note Number is available in Pdf");
     }
      else
      {
        ValidationUtils.verify(false,true,"Credit Note Number is not available in Pdf");
       }
    
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job No").OleValue.toString().trim()))
    {
       if(pdflineSplit[j].includes(jobNumber)){
        Log.Message(jobNumber+" job Number is matching with Pdf");
        ValidationUtils.verify(true,true,jobNumber+": Job Number is matching in pdf");
     
      }
        else
        ValidationUtils.verify(false,true,"JobNumber is not same in pdf");
    }
  }
  
  var index = pdflineSplit.indexOf(clientName);
  if(index>=0)
    ValidationUtils.verify(true,true,clientName+" Client Name is matching in Pdf");
  else
    ValidationUtils.verify(false,true,clientName+" Client Name is not matching in pdf");    
}


