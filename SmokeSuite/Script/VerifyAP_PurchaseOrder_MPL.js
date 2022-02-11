//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

/** 
 * This script implements printing of Purchase order PDF and verify contents against Maconomy
 * @author  : Sai Kiran Vemula
 * @version : 1.0
 * Created Date :07/27/2020
 */
var Project_manager = "";
var Language = "";

/**
  *  This function invokes maconomy and calls subfunctionality methods
  */
  
function verifyPurchaseOrder_MPL(){
TextUtils.writeLog("Verification Of Purchase Order MPL Workspace"); 

//Setting Language in WorkspaceUtils
Language = "";
Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

//Checking Login for Client Creation
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
}

try{
WorkspaceUtils.closeAllWorkspaces();
gotoMenu(); 
verifyAP_PurchaseOrder_MPL();
}
  catch(err){
    Log.Message(err);
  }
}

/**
  *  This function navigates to Purchase Orders under Accounts Payable Workspace
  */

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountsPayable.Exists()){
ImageRepository.ImageSet.AccountsPayable.Click();
}
else{
     ReportUtils.logStep("Fail", "AccountsPayable Workspace not displayed");
     Log.Message("AccountsPayable Workspace not displayed");
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

Log.Message("Language :"+Language)
var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var report_mngmnt;
for(var i=1;i<=childCC;i++){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(report_mngmnt.isVisible()){ 
report_mngmnt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
report_mngmnt.ClickItem("|Purchase Orders");
ReportUtils.logStep_Screenshot();
report_mngmnt.DblClickItem("|Purchase Orders");
}
}
aqUtils.Delay(3000, "Waiting to Load");
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
ReportUtils.logStep("INFO", "Moved to Purchase Orders from Accounts Payable");
Log.Message("Entering Purchase Orders from Accounts Payable");
}
}

/**
  *  This function prints Purchase Orders PDF and validates fields
  */

function verifyAP_PurchaseOrder_MPL(){ 
aqUtils.Delay(2000, "Waiting to Load");
var purchaseOrderTab= Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(purchaseOrderTab);


var allPurchaseOrdersTab = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(allPurchaseOrdersTab);
allPurchaseOrdersTab.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
}
var openAndApprovedPO_Radio = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Open && Approved POs")
Sys.HighlightObject(openAndApprovedPO_Radio);
openAndApprovedPO_Radio.Click();
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
}
var table = Aliases.Maconomy.AccountsPayable.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);

var firstCell = Aliases.Maconomy.AccountsPayable.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.Click();
firstCell.setText(EnvParams.Opco);
if(ImageRepository.ImageSet.LoadedBox.Exists()){}


var closeFilters = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closeFilters.Click();
aqUtils.Delay(1000, "Waiting to Load");

var supplierNo = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
var supplierName = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McValuePickerWidget;
var purchaseOrderNo = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.McTextWidget;
var jobNo = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget2.Composite.McValuePickerWidget;


supplierNo = supplierNo.getText().OleValue.toString().trim();
supplierName= supplierName.getText().OleValue.toString().trim();
purchaseOrderNo= purchaseOrderNo.getText().OleValue.toString().trim();
jobNo= jobNo.getText().OleValue.toString().trim();


var printPurhcaseOrderBtn = Aliases.Maconomy.AP_PurchaseOrder_MPL.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
WorkspaceUtils.waitForObj(printPurhcaseOrderBtn);  
printPurhcaseOrderBtn.HoverMouse();
ReportUtils.logStep_Screenshot("");
printPurhcaseOrderBtn.Click();
TextUtils.writeLog("Print Purchase Order is Clicked and saved"); 
if(ImageRepository.ImageSet.LoadedBox.Exists())
{
}
aqUtils.Delay(5000, Indicator.Text);


var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PurchaseOrder"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PurchaseOrder"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("P_PurchaseOrder")!=-1){
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
ValidationUtils.verify(true,true,"Print Draft Quote is Clicked and PDF is Saved");
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
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Supplier No").OleValue.toString().trim()))
    {
       if(pdflineSplit[j].includes(supplierNo))
       {
        Log.Message(supplierNo+" Supplier Number is matching with Pdf")
        ReportUtils.logStep_Screenshot("");
        ValidationUtils.verify(true,true,supplierNo+": Supplier Number is matching in pdf");
        }
        else
        {
        ValidationUtils.verify(false,true,"Supplier Number is not same in pdf");
        }
    }
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Supplier Name").OleValue.toString().trim()))
    {
       if(pdflineSplit[j].includes(supplierName)){
        Log.Message(supplierName+" Supplier Name is matching with Pdf");
        ValidationUtils.verify(true,true,supplierName+": Supplier Name is matching in pdf");
     
      }
        else
        ValidationUtils.verify(false,true,"Supplier Name is not same in pdf");
    }
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Order No").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(purchaseOrderNo)){
       Log.Message(purchaseOrderNo+" Purchase Order Number is matching with Pdf");
       ValidationUtils.verify(true,true,purchaseOrderNo+": Purchase Order Number is matching in pdf");
     
       }
        else
        ValidationUtils.verify(false,true,"Purchase Order Number is not same in pdf");
      }
          
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job No").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(jobNo)){
        Log.Message(jobNo+" Job Number is matching with Pdf")
        ValidationUtils.verify(true,true,jobNo+": Job Number is matching in pdf");
       break;}
        else{
        ValidationUtils.verify(false,true,"Job Number is not same in pdf");
        break;
        }
      }      
  }
}


