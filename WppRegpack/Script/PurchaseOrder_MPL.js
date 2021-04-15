//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart

/**
 * This script Validate Purchase Order MPL
 * @author  : Sai Kiran Vemula
 * @version : 1.0
 * Created Date :04/15/2021
*/

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName ="PurchaseOrderMPL";


//Global Varibales
var STIME = "";
var company = "";
var poNumber = "";
var Language = "";
var Project_manager = "";
var jobNumber = "";
var noOfPOLines ="";

//Main Function
function PurchaseOrderMPL(){ 
TextUtils.writeLog("Purchase Order MPL Started"); 
Indicator.PushText("waiting for window to open");


//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;


var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4).Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = "GTF Sai Kiran Vemula"//ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}


excelName = EnvParams.path;
workBook = Project.Path+excelName;
STIME = "";
poNumer= "";
company = EnvParams.Opco;
sheetName ="PurchaseOrderMPL";

try{
getDetails();
goToPurchaseOrderMenuItem();
//findPurchaseOrder();
writePOData();
validatePO_PDF();
WorkspaceUtils.closeAllWorkspaces();
}
  catch(err){
    Log.Error(err);
  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4).Click();
WorkspaceUtils.closeAllWorkspaces();
}


//Getting Details to create Sub-Job from Datasheet
function getDetails(){ 

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  poNumber = ReadExcelSheet("PO Number",EnvParams.Opco,"Data Management");
   if((poNumber=="")||(poNumber==null)){
  ValidationUtils.verify(false,true,"Purchase Order Number is needed to validate Purchase Order MPL");
  Runner.Stop();
  }
  
  
  
  
  
}

    

/**
  *  This function Navigates to Purchase Order screen from Accounts Payable workspace
  */
function goToPurchaseOrderMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();

if(ImageRepository.ImageSet.AccountPayable.Exists()){
ImageRepository.ImageSet.AccountPayable.Click();
}
else if(ImageRepository.ImageSet.AccountPayable2.Exists()){
ImageRepository.ImageSet.AccountPayable2.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Purchase Orders").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Purchase Orders from AP Menu");
TextUtils.writeLog("Entering into Purchase Orders from AP Menu");
}


//Validating Working Estimate is Approved or Not
function findPurchaseOrder(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
 
  var myPos = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "My POs").OleValue.toString().trim())
  WorkspaceUtils.waitForObj(myPos);
  Sys.HighlightObject(myPos);
  myPos.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
 
  var companyNo = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget
  WorkspaceUtils.waitForObj(companyNo);
  companyNo.Click();
  companyNo.Keys("[Tab]");

  var purchaseNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox
  WorkspaceUtils.waitForObj(purchaseNo);
  purchaseNo.Click();
  purchaseNo.setText(poNumber);

  var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
  WorkspaceUtils.waitForObj(table);

  while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ }
    
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==poNumber){ 
  flag=true;    
  break;
  
  }
  else{ 
  table.Keys("[Down]");
  }
  }
  if(!flag)
  {
    ValidationUtils.verify(flag,true,"Purchase Order is not available in system");
     TextUtils.writeLog("Purchase Order is available not in system");
  }
  
   if(flag){
   ValidationUtils.verify(flag,true,"Purchase Order is available in system");
  TextUtils.writeLog("Purchase Order is available in system");
   var closefilter =Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter
  WorkspaceUtils.waitForObj(closefilter);
  closefilter.HoverMouse();
  ReportUtils.logStep_Screenshot();
  closefilter.Click();
 }
 
 }


// Write Purchase Order Data to Excel
function writePOData(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){  }
   

var specification = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var poMPL = "PurchaseOrderMPL";

var q = 0;
poDetails = [];
noOfPOLines = specification.getItemCount()
Log.Message(specification.getItemCount())
for(var i=0;i<noOfPOLines;i++){ 
Log.Message("i: "+i);
  var job_no = specification.getItem(i).getText_2(1).OleValue.toString().trim();  
  var description = specification.getItem(i).getText_2(4).OleValue.toString().trim();
  var quantity = specification.getItem(i).getText_2(5).OleValue.toString().trim();
  var unitPrice = specification.getItem(i).getText_2(6).OleValue.toString().trim();
  var amount = specification.getItem(i).getText_2(7).OleValue.toString().trim();
  var total = specification.getItem(i).getText_2(8).OleValue.toString().trim();
  poDetails[q] = job_no+"*"+description+"*"+quantity+"*"+unitPrice+"*"+amount+"*"+total;
  Log.Message(poDetails[q]);
  q++;
  ExcelUtils.setExcelName(workBook,poMPL, true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"JobNumber",poMPL,job_no);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+q,poMPL,description);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+q,poMPL,quantity);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"UnitPrice_"+q,poMPL,unitPrice);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Amount_"+q,poMPL,amount);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Total_"+q,poMPL,total);
  }
Log.Message(q)


var print = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
WorkspaceUtils.waitForObj(print);    
print.HoverMouse();
ReportUtils.logStep_Screenshot("");
print.Click();
TextUtils.writeLog("Print PO is Clicked and saved"); 
aqUtils.Delay(5000, Indicator.Text);
 

var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PurchaseOrder"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "P_PurchaseOrder"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("P_PurchaseOrder")!=-1){
    aqUtils.Delay(7000, Indicator.Text);
WorkspaceUtils.waitForObj(pdf);
Sys.HighlightObject(pdf)
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x41); //A 
Sys.Desktop.KeyUp(0x12); 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x41);
    
if(ImageRepository.PDF.ChooseFolder.Exists())
ImageRepository.PDF.ChooseFolder.Click();
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
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
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
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PurchaseOrderMPL",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")

}


function validatePO_PDF()
{
  
  var fileName = "";
  var Language = "";
  Language = EnvParams.LanChange(EnvParams.Language);
  WorkspaceUtils.Language = Language;  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  
  fileName = ExcelUtils.getRowDatas("PurchaseOrderMPL",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"PurchaseOrderMPL PDF is needed to validate");
  }
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName);
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
  var sheetName = "PurchaseOrderMPL";
  var pdflineSplit = docObj.split("\r\n");

   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "PURCHASE ORDER").OleValue.toString().trim());
        if(index>=0){
          ReportUtils.logStep("INFO","Heading is available Pdf")
          ValidationUtils.verify(true,true,"Heading is available Pdf")
          TextUtils.writeLog("Heading is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Heading is not available Pdf") 
 
   index = docObj.indexOf(poNumber); 
   if(index>=0){
          ReportUtils.logStep("INFO",poNumber+ " PONumber is matching with Pdf")
          ValidationUtils.verify(true,true,poNumber + " PONumber is matching with Pdf")
          TextUtils.writeLog(poNumber+" PONumber is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"PONumber is not same in Pdf");
                   
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var temp=0;
  for(var i=1;i<5;i++){
  jobNumber = ExcelUtils.getColumnDatas("JobNumber",EnvParams.Opco);
  var total = ExcelUtils.getColumnDatas("Total_"+i,EnvParams.Opco);
  if (total!=""){
  total = parseFloat(total.replace(/,/g, ''));
 // Log.Message(total)
  }
  if(total>0){
  temp = temp + total;
 // Log.Message(temp)       
  }
  }    
 var exclTaxTotal = temp.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); 
    Log.Message(exclTaxTotal)
    index = docObj.indexOf(exclTaxTotal);          
    if(index>=0){
          ReportUtils.logStep("INFO",exclTaxTotal +" Total Excluding Tax is matching with Pdf")
          ValidationUtils.verify(true,true,exclTaxTotal+" Total Excluding Tax is matching with Pdf")
          TextUtils.writeLog(exclTaxTotal+" Total Excluding Tax is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,exclTaxTotal+"Total Excluding Tax is not same in Pdf");      
    
   index = docObj.indexOf(jobNumber);          
    if(index>=0){
          ReportUtils.logStep("INFO",jobNumber+" JobNumber is matching with Pdf")
          ValidationUtils.verify(true,true,jobNumber+" JobNumber is matching with Pdf")
          TextUtils.writeLog(jobNumber+" JobNumber is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,jobNumber+"JobNumber is not same in Pdf");            
}
 
 

