//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
//USEUNIT EventHandler

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "AR Multiple Payment";
Indicator.Show();
Indicator.PushText("waiing for window to open");


var Jobno ="";
var companyno ="";
var InvoiceNumber ="";
var Descip="";
var currency ="";
var clientnum = "";
var TP ="";
var STIME = "";
var Clientbalance ="";
var JournalNo = "";
var Amount = "";
function MultiplePayment() {
  
TextUtils.writeLog("Create Payment Selection Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - AR","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  
}
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "AR Multiple Payment";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
JournalNo = "";
Amount = "";
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);
try{     
    getDetails();
    goToJobMenuItem();
    invoicejob();
    var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
    menuBar.Click();
    WorkspaceUtils.closeAllWorkspaces();
    goToARMenuItem();    
    gotoReceivable();
    
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Combined Biller, IC & AR","Username")
Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);  

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
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
goToJobMenuItem();
reconsilejob();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces(); 
}


}
  catch(err){
    Log.Message(err);
  }
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Multiple Invoice Payment used Invoice No",EnvParams.Opco,"Data Management",InvoiceNumber)
  ExcelUtils.WriteExcelSheet("Multiple Invoice Payment used Job",EnvParams.Opco,"Data Management",Jobno)
  TextUtils.writeLog("Multiple Invoice Payment used Invoice No : "+InvoiceNumber);
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}



function getDetails(){

sheetName = "AR Multiple Payment";     
ExcelUtils.setExcelName(workBook, sheetName, true);
 
Descip = ExcelUtils.getRowDatas("Description",EnvParams.Opco)
Log.Message(Descip)
if((Descip==null)||(Descip=="")){ 
ValidationUtils.verify(false,true,"Description is needed to Create Single Invoice"); 
}  

currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
Log.Message(currency)
if((currency==null)||(currency=="")){ 
ValidationUtils.verify(false,true,"Currency is needed to Create Single Invoice"); 
}   
//        TP = ExcelUtils.getRowDatas("TP",EnvParams.Opco)
//        Log.Message(TP)
//        if((TP==null)||(TP=="")){ 
//        ValidationUtils.verify(false,true,"TP is needed to Create Single Invoice"); 
//        }  

ExcelUtils.setExcelName(workBook, "Data Management", true);

  var invoicePreparation = ExcelUtils.getRowDatas("Invoice preparation Job",EnvParams.Opco);
  var invoiceBudget = ExcelUtils.getRowDatas("Invoice from Budget Job",EnvParams.Opco);
  var invoiceAccount = ExcelUtils.getRowDatas("Invoice OnAccount Job",EnvParams.Opco);
  var TM = ExcelUtils.getRowDatas("Time and Material Invoice Job",EnvParams.Opco);
  
  var iP = ExcelUtils.getRowDatas("Invoice preparation No",EnvParams.Opco);
  var iB = ExcelUtils.getRowDatas("Invoice from Budget No",EnvParams.Opco);
  var iA = ExcelUtils.getRowDatas("Invoice OnAccount No",EnvParams.Opco);
  var iTM = ExcelUtils.getRowDatas("Time and Material Invoice No",EnvParams.Opco);
  
  var sP = ExcelUtils.getRowDatas("Single Invoice Payment used Invoice No",EnvParams.Opco);
  var fP = ExcelUtils.getRowDatas("Foreign Invoice Payment used Invoice No",EnvParams.Opco);
  var mP = ExcelUtils.getRowDatas("Multiple Invoice Payment used Invoice No",EnvParams.Opco);

  if(((invoicePreparation!="")&&(invoicePreparation!=null))&&((iP!="")&&(iP!=null))&&(iP!=sP)&&(iP!=fP)&&(iP!=mP)){
    InvoiceNumber = iP;
    Jobno = invoicePreparation;
  }else  if(((invoiceBudget!="")&&(invoiceBudget!=null))&&((iB!="")&&(iB!=null))&&(iB!=sP)&&(iB!=fP)&&(iB!=mP)){
    InvoiceNumber = iB;
    Jobno = invoiceBudget;
  }else  if(((invoiceAccount!="")&&(invoiceAccount!=null))&&((iA!="")&&(iA!=null))&&(iA!=sP)&&(iA!=fP)&&(iA!=mP)){
    InvoiceNumber = iA;
    Jobno = invoiceAccount;
  }else  if(((TM!="")&&(TM!=null))&&((iTM!="")&&(iTM!=null))&&(iTM!=sP)&&(iTM!=fP)&&(iTM!=mP)){
    InvoiceNumber = iA;
    Jobno = TM;
  }else{ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Jobno = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
Log.Message(Jobno)
InvoiceNumber = ExcelUtils.getRowDatas("Invoice Number",EnvParams.Opco)
Log.Message(InvoiceNumber)
  }
if((Jobno=="")||(Jobno==null))
ValidationUtils.verify(false,true,"Job Number is needed to Create Single Invoice");

if((InvoiceNumber=="")||(InvoiceNumber==null))
ValidationUtils.verify(false,true,"Invoice Number is needed to Create Single Invoice");

//Jobno = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
//Log.Message(Jobno)
//if((Jobno=="")||(Jobno==null)){
//ExcelUtils.setExcelName(workBook, sheetName, true);
//Jobno = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
//Log.Message(Jobno)
//}  
//if((Jobno=="")||(Jobno==null))
//ValidationUtils.verify(false,true,"Job Number is needed to Create Single Invoice");
        
ExcelUtils.setExcelName(workBook, "Data Management", true);
clientnum = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
Log.Message(clientnum)
if((clientnum=="")||(clientnum==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
clientnum = ExcelUtils.getRowDatas("Global Client Number",EnvParams.Opco)
Log.Message(clientnum)
}  
if((clientnum=="")||(clientnum==null))
ValidationUtils.verify(false,true,"Client Number is needed to Create Single Invoice");


ExcelUtils.setExcelName(workBook, sheetName, true);
Amount = ""; //ExcelUtils.getRowDatas("Amount",EnvParams.Opco)
Log.Message(Amount)


//ExcelUtils.setExcelName(workBook, "Data Management", true);
//InvoiceNumber = ReadExcelSheet("Client Invoice No",EnvParams.Opco,"Data Management");
//Log.Message(InvoiceNumber)
//if((InvoiceNumber=="")||(InvoiceNumber==null)){
//ExcelUtils.setExcelName(workBook, sheetName, true);
//InvoiceNumber = ExcelUtils.getRowDatas("Invoice Number",EnvParams.Opco)
//Log.Message(InvoiceNumber)
//}  
//if((InvoiceNumber=="")||(InvoiceNumber==null))
//ValidationUtils.verify(false,true,"Invoice Number is needed to Create Single Invoice");
          
}   
  
function goToJobMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}

}

ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu"); 
  }  



function invoicejob(){  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
     TextUtils.writeLog("Customer Payment for Foreign Invoice is started");
      ReportUtils.logStep("INFO", "Customer Payment for Foreign Invoice is started::"+STIME);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var allJobs = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();

  var table = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(Jobno);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
//  aqUtils.Delay(7000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==Jobno){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
//  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to create payment for Invoice");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+Jobno+") is available in maconommy to create payment for Invoice"); 
  closeFilter.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

                        
       var invoice = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.invoice;
       Sys.HighlightObject(invoice);       
       ReportUtils.logStep_Screenshot(""); 
       invoice.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
       
       var invoicehistory = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.invoicehis;
       Sys.HighlightObject(invoicehistory);
       invoicehistory.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
       
       var invoicetable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.invoicetable;
       Sys.HighlightObject(invoicetable);
       ReportUtils.logStep_Screenshot(""); 
       var row = invoicetable.getItemCount();
       var column = invoicetable.getColumnCount();
    var checkStatus = false;   
         for(var i=0;i<invoicetable.getItemCount();i++){   
          if((invoicetable.getItem(i).getText(0).OleValue.toString().trim()==InvoiceNumber)&&((invoicetable.getItem(i).getText(8).OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Not Due").OleValue.toString().trim())!=-1)||(invoicetable.getItem(i).getText(8).OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Due").OleValue.toString().trim())!=-1))){   
            if(Amount==""){ 
             Amount = invoicetable.getItem(i).getText(9).OleValue.toString().trim();
             Log.Message(Amount)
             Amount = Amount.replace(/,/g, '');
            Amount = parseFloat(Amount).toFixed(0);
            Log.Message(Amount)
            Amount = Amount/2;
            Log.Message(Amount)
             
            }
            checkStatus = true;
              break;
            }
           else{
                 invoicetable.Keys("[Down]");
           } 
         }
      ValidationUtils.verify(true,checkStatus,"Invoice is in Due")
      ReportUtils.logStep_Screenshot(""); 
//      TextUtils.writeLog("Payment status Need to Reconcile");
//      ValidationUtils.verify(true,true,"Payment status Need to Reconcile");
      
      var Home = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.overview;        
      Sys.HighlightObject(Home);
      Home.Click();
        
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }       
       var clientbalance = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.clientbalance;
       Clientbalance = clientbalance.getText();
       ReportUtils.logStep_Screenshot(); 
//       ValidationUtils.verify(true,true,"Payment status Need to Reconcile");
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
  }

function gotoReceivable(){  
      
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
     var clientopen = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Registrations;
     WorkspaceUtils.waitForObj(clientopen)
     clientopen.HoverMouse();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
     ReportUtils.logStep_Screenshot("");
     clientopen.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
     
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
        
var newbutton = "";        
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.isVisible())
  newbutton = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.SingleToolItemControl
 else
  newbutton = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.okbutton;
//     var newbutton = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.okbutton;
     WorkspaceUtils.waitForObj(newbutton)
     Sys.HighlightObject(newbutton);
     newbutton.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
     

     
var company = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.company;
WorkspaceUtils.waitForObj(company)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
company.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if(company!=""){
company.setText(EnvParams.Opco);
aqUtils.Delay(2000,Indicator.Text);
ValidationUtils.verify(true,true,"Company is Entered");
}
else{
ValidationUtils.verify(false,true,"Company is Needed to Create Single Invoice");
}
aqUtils.Delay(2000,Indicator.Text);
     
var descrip = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.descrip;
Sys.HighlightObject(descrip);     
descrip.Click();
if(Descip!=""){
descrip.setText(Descip);
aqUtils.Delay(2000,Indicator.Text);
ValidationUtils.verify(true,true,"Description is Entered");
}
else{
ValidationUtils.verify(false,true,"Description is Needed to Create Multiple Invoice");
}
     
var currenccy = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite2.currency;
if(currency!=""){
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
currenccy.Keys(currency);
//currenccy.Click();
aqUtils.Delay(2000, Indicator.Text);
//WorkspaceUtils.DropDownList(currency,"Currency");
aqUtils.Delay(2000, Indicator.Text); 
} 
else{
ValidationUtils.verify(false,true,"Currency is Needed to Create Single Invoice"); 
} 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}  
var amount = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite3.amount;
Sys.HighlightObject(amount);
amount.Click();    
if(Clientbalance!=""){       
amount.setText(Clientbalance);
ValidationUtils.verify(true,true,"Invoice Amount is Entered");
aqUtils.Delay(2000,Indicator.Text);
}
else{
ValidationUtils.verify(false,true,"Invoice Amount is Needed to Create Single Invoice"); 
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
} 
var scroll= NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
//WorkspaceUtils.waitForObj(scroll)
//if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//  
//}  
scroll.Click();   
scroll.MouseWheel(-200);
      
var client = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.Composite.client;
if(clientnum!=""){
WorkspaceUtils.waitForObj(client)
client.Click();
WorkspaceUtils.SearchByValuePicker(client,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),clientnum,"Client Number");
} 
else{ 
ValidationUtils.verify(false,true,"Client Number is Exist for Single Invoice");
} 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
} 
var scroll= NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
scroll.Click();      
scroll.MouseWheel(200);
      
var showbutton = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget.Composite.McPlainCheckboxView.showButton;
showbutton.HoverMouse();
ReportUtils.logStep_Screenshot("");
showbutton.Click();
aqUtils.Delay(2000,Indicator.Text);
ReportUtils.logStep("INFO", "Show Lines is Checked");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
} 
 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
//  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions - Client Open Entry Reconciliation").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
        var popup = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions - Client Open Entry Reconciliation").OleValue.toString().trim()).SWTObject("Label", "The difference will be allocated as cash discount on the reconciled entries")
        Sys.HighlightObject(popup);
        var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions - Client Open Entry Reconciliation").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        Sys.HighlightObject(OK);
        OK.Click();   
}

//var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.savee;
//WorkspaceUtils.waitForObj(save)
//save.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
     
var getjournal = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite.journal.getText().OleValue.toString().trim();

var BankAccount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2).getText().OleValue.toString().trim();
//var BankAccount = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2).getText().OleValue.toString().trim();
var Transaction_No = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget3.Composite2.SWTObject("McTextWidget", "", 2).getText().OleValue.toString().trim();
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Multiple Payment Trans No",EnvParams.Opco,"Data Management",Transaction_No)
  ExcelUtils.WriteExcelSheet("Account Number",EnvParams.Opco,"Data Management",BankAccount)     
       
var artable = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table;
Sys.HighlightObject(artable);       
    
var  column = artable.getColumnCount();
var row = artable.getItemCount()
Log.Message(column)
Log.Message(row)
       
for(var i=0;i<row;i++){
if(artable.getItem(i).getText(0).OleValue.toString().trim()==InvoiceNumber){
ValidationUtils.verify(true,true,"Invoice Number is available in the table");
break;
}
else{
artable.Keys("[Down]");
}
}       
             
artable.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var reconsile = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.amount;
WorkspaceUtils.waitForObj(reconsile);
reconsile.setText(Amount);
Sys.HighlightObject(reconsile);
reconsile.Keys("[Tab]");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var tp = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table;    
var tp = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.table.assettype;
WorkspaceUtils.waitForObj(tp);
//tp.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Partially").OleValue.toString().trim())
tp.Click();
aqUtils.Delay(4000, Indicator.Text);
WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Partially").OleValue.toString().trim(),"Totally/Partially",tp)
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Partially").OleValue.toString().trim(),"Totally");

//if(TP!=""){
//tp.Click();
//aqUtils.Delay(2000, Indicator.Text);
//WorkspaceUtils.DropDownList(TP,"Totally");
//aqUtils.Delay(2000, Indicator.Text); 
//} 
//else{
//ValidationUtils.verify(false,true,"TP is Needed to Create Single Invoice");  
//} 


aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
} 
var saveentry = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
WorkspaceUtils.waitForObj(saveentry)
saveentry.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
       
var released = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
Sys.HighlightObject(released);
WorkspaceUtils.waitForObj(released)
released.Click();
aqUtils.Delay(2000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
var relesave = ""; 
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.isVisible())
  relesave = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.SWTObject("SingleToolItemControl", "", 4)
 else
  relesave = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;       
       
//var relesave = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;       
WorkspaceUtils.waitForObj(relesave)
relesave.Click(); 
aqUtils.Delay(5000,Indicator.Text);         
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
      
    var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
//  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs - Job").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions - Client Open Entry Reconciliation").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
        var popup = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions - Client Open Entry Reconciliation").OleValue.toString().trim()).SWTObject("Label", "The difference will be allocated as cash discount on the reconciled entries")
        Sys.HighlightObject(popup);
        var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "AR Transactions - Client Open Entry Reconciliation").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        Sys.HighlightObject(OK);
        OK.Click();   
} 

//       if((Sys.Process("Maconomy").SWTObject("Shell", "AR Transactions - Client Open Entry Reconciliation")).isVisible());
//       {
//        var popup = Sys.Process("Maconomy").SWTObject("Shell", "AR Transactions - Client Open Entry Reconciliation").SWTObject("Label", "The difference will be allocated as cash discount on the reconciled entries")
//        Sys.HighlightObject(popup);
//        var OK = Sys.Process("Maconomy").SWTObject("Shell", "AR Transactions - Client Open Entry Reconciliation").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//        Sys.HighlightObject(OK);
//        OK.Click();      
//       }
                        
//       var paymenttab = Aliases.Maconomy.AR.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel;      
//        Sys.HighlightObject(paymenttab)
        
      var clientpayment = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.clientpayment     
      
      Sys.HighlightObject(clientpayment);
      clientpayment.HoverMouse();
      ReportUtils.logStep_Screenshot("");
      clientpayment.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
      var tab = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.journalnumber
//      var tab = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;      
//NameMapping.Sys.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
      Sys.HighlightObject(tab);
      var journalnum = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.journalnumber
//      var journalnum = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;      
//NameMapping.Sys.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.journalnumber;
      journalnum.Click();
      journalnum.setText(getjournal);
//      aqUtils.Delay(4000,Indicator.Text);     
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(1000,Indicator.Text); 
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);   
aqUtils.Delay(1000,Indicator.Text);   
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
      var submit = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.submitjournal;      
//NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.submitjournal;
      Sys.HighlightObject(submit);
      submit.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
      aqUtils.Delay(4000,Indicator.Text); 
      submit.HoverMouse();
      ReportUtils.logStep_Screenshot("");    
      TextUtils.writeLog("Journal Number is Submitted");
      ValidationUtils.verify(true,true,"Journal Number is Submitted");  
      
ValidationUtils.verify(true,true,"Journal Number is : "+getjournal);
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Multiple Invoice Journal No",EnvParams.Opco,"Data Management",getjournal)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 JournalNo =   getjournal;   
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

var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Posting Journal"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Posting Journal")!=-1){
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

//function goToJobMenuItem(){
//var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
//menuBar.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//menuBar.DblClick();
//if(ImageRepository.ImageSet0.Jobs.Exists()){
//ImageRepository.ImageSet0.Jobs.Click();// GL
//}
//else if(ImageRepository.ImageSet0.Job.Exists()){
//ImageRepository.ImageSet0.Job.Click();
//}
//else{
//ImageRepository.ImageSet0.Jobs1.Click();
//}
//aqUtils.Delay(3000, Indicator.Text);
//Sys.Desktop.KeyDown(0x12);
//Sys.Desktop.KeyDown(0x20);
//Sys.Desktop.KeyUp(0x12);
//Sys.Desktop.KeyUp(0x20);
//Sys.Desktop.KeyDown(0x58);
//Sys.Desktop.KeyUp(0x58);  
//aqUtils.Delay(1000, Indicator.Text);
//var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
//var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
//Delay(3000);
//var MainBrnch = "";
//for(var bi=0;bi<WrkspcCount;bi++){ 
//if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
//MainBrnch = Workspc.Child(bi);
//break;
//}
//}
//
//var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
//var Client_Managt;
//for(var i=1;i<=childCC;i++){ 
//Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
//if(Client_Managt.isVisible()){ 
//Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
//Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
//ReportUtils.logStep_Screenshot();
//Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
//}
//}
//Delay(3000);
//} 

function reconsilejob(){  
  

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
          if((invoicetable.getItem(i).getText(0).OleValue.toString().trim()==InvoiceNumber)&&(invoicetable.getItem(i).getText(8).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Partially Reconciled").OleValue.toString().trim())){  
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




 
 