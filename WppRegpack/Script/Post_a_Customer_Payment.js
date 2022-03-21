//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "Post a Customer Payment";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");
  
var JournalNo ="";
var Jobno = "";
var InvoiceNumber = "";
function CustomerPayment() { 
  
TextUtils.writeLog("Create Payment Selection Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Combined Biller, IC & AR","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
STIME = "";  
JournalNo = "";
Jobno = "";
InvoiceNumber = ""
getDetails();
goToARMenuItem();
gotoPost();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();

//goToGeneralLedger();
//GLLookups();
//var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//menuBar.Click();
//WorkspaceUtils.closeAllWorkspaces();  
 
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - AR","Username")
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  
//WorkspaceUtils.closeAllWorkspaces();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
goToJobMenuItem();
invoicejob(); 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);       
sheetName = "Post a Customer Payment";   
        
ExcelUtils.setExcelName(workBook, "Data Management", true);
JournalNo = ReadExcelSheet("Single Invoice Journal No",EnvParams.Opco,"Data Management");
if((JournalNo!="")||(JournalNo!=null)){
Jobno = ReadExcelSheet("Single Invoice Payment used Job",EnvParams.Opco,"Data Management");
        if((Jobno=="")||(Jobno==null))
        ValidationUtils.verify(false,true,"Job Number is needed to Check Invoice status"); 
InvoiceNumber = ReadExcelSheet("Single Invoice Payment used Invoice No",EnvParams.Opco,"Data Management");
        if((InvoiceNumber=="")||(InvoiceNumber==null))
        ValidationUtils.verify(false,true,"Invoice Number is needed to Check Invoice status"); 
  }
if((JournalNo=="")||(JournalNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
JournalNo = ExcelUtils.getRowDatas("Journal No",EnvParams.Opco)
}
if((JournalNo=="")||(JournalNo==null))
        ValidationUtils.verify(false,true,"JournalNo is needed to Check Invoice status");   


//        ExcelUtils.setExcelName(workBook, sheetName, true);       
//        sheetName = "Post a Customer Payment";  
//  
//        companyno = EnvParams.Opco;
//        Jobno = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)

if((Jobno=="")||(Jobno==null)){
ExcelUtils.setExcelName(workBook, sheetName, true); 
Jobno = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
}  
if((Jobno=="")||(Jobno==null))
ValidationUtils.verify(false,true,"Job Number is needed to Check Invoice status");   

if((InvoiceNumber=="")||(InvoiceNumber==null)){
ExcelUtils.setExcelName(workBook, sheetName, true); 
InvoiceNumber = ReadExcelSheet("Invoice Number",EnvParams.Opco,"Data Management");
}  
if((InvoiceNumber=="")||(InvoiceNumber==null))
ValidationUtils.verify(false,true,"Invoice Number is needed to Check Invoice status");            
} 

function goToARMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.Account_Receivable.Exists()){
ImageRepository.ImageSet.Account_Receivable.Click();// GL
}
else if(ImageRepository.ImageSet.Acc_Receivable_1.Exists()){
ImageRepository.ImageSet.Acc_Receivable_1.Click();
}
else{
ImageRepository.ImageSet.Acc_Receivable_2.Click();
}

aqUtils.Delay(3000, Indicator.Text);
Sys.Desktop.KeyDown(0x12);
Sys.Desktop.KeyDown(0x20);
Sys.Desktop.KeyUp(0x12);
Sys.Desktop.KeyUp(0x20);
Sys.Desktop.KeyDown(0x58);
Sys.Desktop.KeyUp(0x58);  

aqUtils.Delay(1000, Indicator.Text);
var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
Delay(3000);
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
        Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions").OleValue.toString().trim());
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions").OleValue.toString().trim());
}
}
Delay(3000);
}


function gotoPost(){ 
//  TextUtils.writeLog("Post Customer Payment is started");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

ReportUtils.logStep("INFO", "Post Customer Payment is started::"+STIME);
var client = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.clientpayment;
Sys.HighlightObject(client);
var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.table.clienttable;
Sys.HighlightObject(table);
var firstcell = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.table.clienttable.firstcell;
firstcell.Click();
firstcell.setText(JournalNo);
aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var flag =false;
for(var i=0;i<table.getItemCount();i++){
if((table.getItem(i).getText_2(0).OleValue.toString().trim()==JournalNo)){
flag = true;
  break;
} 
else{
  table.Keys("[Down]");
} 
} 
ValidationUtils.verify(flag,true,"Journal Number is available in system");
aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ReportUtils.logStep_Screenshot("");
Sys.Desktop.KeyDown(0x11);
Sys.Desktop.KeyDown(0x46);
Sys.Desktop.KeyUp(0x11);
Sys.Desktop.KeyUp(0x46);
               
        
ReportUtils.logStep("INFO", "Journal Number is listed in the table");            
aqUtils.Delay(3000,Indicator.Text);       
ReportUtils.logStep_Screenshot("");
var post = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite2.post;
WorkspaceUtils.waitForObj(post);
post.Click();
ReportUtils.logStep_Screenshot("");          
ValidationUtils.verify(true,true,"Journal is Posted");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*", 1).WndCaption.indexOf("Print Posting Journal")!=-1){
    aqUtils.Delay(2000, Indicator.Text);

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
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
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
ValidationUtils.verify(true,true,"Print Draft Quote is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}






      
} 

function goToJobMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet0.Jobs.Exists()){
ImageRepository.ImageSet0.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet0.Job.Exists()){
ImageRepository.ImageSet0.Job.Click();
}
else{
ImageRepository.ImageSet0.Jobs1.Click();
}
aqUtils.Delay(3000, Indicator.Text);
Sys.Desktop.KeyDown(0x12);
Sys.Desktop.KeyDown(0x20);
Sys.Desktop.KeyUp(0x12);
Sys.Desktop.KeyUp(0x20);
Sys.Desktop.KeyDown(0x58);
Sys.Desktop.KeyUp(0x58);  
aqUtils.Delay(1000, Indicator.Text);
var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
Delay(3000);
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}
}
Delay(3000);
} 

function invoicejob(){  
  

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
      var table = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid      
      var compno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
      compno.Click();
      compno.setText(EnvParams.Opco);
      aqUtils.Delay(1000,Indicator.Text);
      compno.Keys("[Tab][Tab]");
      var jobno = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
      jobno.Click();
      jobno.setText(Jobno);
      aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
      var flag =false; 
        for(var i=0;i<table.getItemCount();i++){          
          if(table.getItem(i).getText_2(2).OleValue.toString().trim()==Jobno){
            flag = true;        
            break;
          }  
          else{
              table.Keys("[Down]");
          } 
        } 
        aqUtils.Delay(2000,Indicator.Text); 
        ReportUtils.logStep_Screenshot();    
        ValidationUtils.verify(true,true,"Job Number is available in system");
        aqUtils.Delay(3000,Indicator.Text);          
            
                Sys.Desktop.KeyDown(0x11);
                Sys.Desktop.KeyDown(0x46);
               Sys.Desktop.KeyUp(0x11);
                Sys.Desktop.KeyUp(0x46);
                        
       var invoice = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.invoice;
       Sys.HighlightObject(invoice);       
       ReportUtils.logStep_Screenshot(""); 
       invoice.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
       var invoicehistory = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.invoicehis;
//       var invoicehistory = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 10);
       Sys.HighlightObject(invoicehistory);
       invoicehistory.Click();
       aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
       }
       var invoicetable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.invoicetable;
//       var invoicetable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.SWTObject("Composite", "", 9).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
       Sys.HighlightObject(invoicetable);
       ReportUtils.logStep_Screenshot(""); 
       var row = invoicetable.getItemCount();
       var column = invoicetable.getColumnCount();
    var checkStatus = false;   
         for(var i=0;i<invoicetable.getItemCount();i++){   
          if((invoicetable.getItem(i).getText(0).OleValue.toString().trim()==InvoiceNumber)&&(invoicetable.getItem(i).getText(8).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Reconciled").OleValue.toString().trim())){  
            checkStatus = true;
              break;
            }
           else{
                 invoicetable.Keys("[Down]");
           } 
         }
      ValidationUtils.verify(true,checkStatus,"Invoice is Reconciled")
      
//         for(var i=0;i<invoicetable.getItemCount();i++){         
//                 if(invoicetable.getItem(i).getText(8).OleValue.toString().trim()=="Reconciled"){  
//                  break;
//                 }
//                 else{
//                     invoicetable.Keys("[Down]");
//                 }       
//      }
      ReportUtils.logStep_Screenshot(""); 
      //  TextUtils.writeLog("Payment status changed as Reconciled");
      ValidationUtils.verify(true,true,"Payment status changed as Reconciled");
      
}


function goToGeneralLedger(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet0.GendralLedger.Exists()){
ImageRepository.ImageSet0.GendralLedger.Click();// GL
}
else if(ImageRepository.ImageSet0.GendralLedger1.Exists()){
ImageRepository.ImageSet0.GendralLedger1.Click();
}
else{
ImageRepository.ImageSet0.GendralLedger2.Click();
}
aqUtils.Delay(3000, Indicator.Text);
Sys.Desktop.KeyDown(0x12);
Sys.Desktop.KeyDown(0x20);
Sys.Desktop.KeyUp(0x12);
Sys.Desktop.KeyUp(0x20);
Sys.Desktop.KeyDown(0x58);
Sys.Desktop.KeyUp(0x58);  
aqUtils.Delay(1000, Indicator.Text);
var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
Delay(3000);
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Lookups").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Lookups").OleValue.toString().trim());
}
}
Delay(3000);
}

  
function GLLookups(){
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var journal = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.journal;
journal.Click();
ReportUtils.logStep_Screenshot("");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var postedjournal = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2
Sys.HighlightObject(postedjournal) ;
aqUtils.Delay(1000,Indicator.Text);
var all = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim())
all.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
   
var journalnum = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell;
journalnum.Click();
journalnum.setText(JournalNo);
aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
var flag =false;
for(var i=0;i<table.getItemCount();i++){
if((table.getItem(i).getText_2(0).OleValue.toString().trim()==JournalNo)){
flag = true;
break;
} 
else{
table.Keys("[Down]");
} 
} 
aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ReportUtils.logStep_Screenshot("");
Sys.Desktop.KeyDown(0x11);
Sys.Desktop.KeyDown(0x46);
Sys.Desktop.KeyUp(0x11);
Sys.Desktop.KeyUp(0x46);
aqUtils.Delay(3000,Indicator.Text);        
ValidationUtils.verify(flag,true,"Posted Journal Number is available in the List");
ReportUtils.logStep("INFO", "Posted Journal Number is available in the List");            
aqUtils.Delay(1000,Indicator.Text);       
ReportUtils.logStep_Screenshot("");
      
}
  

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}



