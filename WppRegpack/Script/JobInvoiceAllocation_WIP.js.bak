//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "JobInvoiceAllocation_WIP";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var jobNumber,EmpNo = "";
var Estimatelines = [];
var LatestTran = ""
var Language = "";
//Main Function
function InvoiceAllocation(){ 
TextUtils.writeLog("Job Invoice Allocation (Without WIP) Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Biller",EnvParams.Opco);
//if((Project_manager=="")||(Project_manager==null))
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if((Project_manager=="")||(Project_manager==null))
ValidationUtils.verify(false,true,"Login Credentials required for anyone of Agency - Biller or Agency - Finance,");

Log.Message(Project_manager);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
 
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "JobInvoiceAllocation_WIP";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
jobNumber,EmpNo,LatestTran = "";
Estimatelines = [];

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Job Invoice Allocation (Without WIP) started::"+STIME);
getDetails();
gotoMenu();
gotoAllocation();
WorkspaceUtils.closeAllWorkspaces();
gotoGeneralJournal();
GlLookups()

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){ 
sheetName ="JobInvoiceAllocation_WIP";  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  jobTemplate = ReadExcelSheet("Job Template",EnvParams.Opco,"Data Management");
  Log.Message((jobNumber=="")||(jobNumber==null))
  Log.Message(jobTemplate.indexOf("FP")==-1)
  Log.Message(((jobNumber=="")||(jobNumber==null))||(jobTemplate.indexOf("FP")==-1))
  if(((jobNumber=="")||(jobNumber==null))||(jobTemplate.indexOf("FP")==-1)){
  jobNumber = ReadExcelSheet("Invoice preparation Job",EnvParams.Opco,"Data Management"); 
    }
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed for Job Invoice Allocation (Without WIP)");
  Log.Message(jobNumber)
//  EmpNo = ReadExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management");
//  if((EmpNo=="")||(EmpNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  EmpNo = ExcelUtils.getRowDatas("Employee Number",EnvParams.Opco)
//  }
  if((EmpNo=="")||(EmpNo==null))
  ValidationUtils.verify(false,true,"Employee Number is needed for Job Invoice Allocation (Without WIP)");
  
}


function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Jobs.Exists()){
ImageRepository.ImageSet.Jobs.Click();// GL
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


ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}

//function gotoAllocation(){ 
//var allJobs = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
//allJobs.Click();
//
//
//var labels = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
//WorkspaceUtils.waitForObj(labels);
//for(var i=0;i<labels.ChildCount;i++){ 
//  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf("Now showing")!=-1)){
//    labels = labels.Child(i);
//    break;
//  }
//}
//
//WorkspaceUtils.waitForObj(labels);
//
//  var table = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//  var firstcell = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
//  var closeFilter = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//  WorkspaceUtils.waitForObj(firstcell);
//  firstcell.forceFocus();
//  firstcell.setVisible(true);
//  firstcell.ClickM();
//  Sys.Desktop.KeyDown(0x09); // Press Ctrl
//  aqUtils.Delay(1000, Indicator.Text);
//  Sys.Desktop.KeyDown(0x09);
//  aqUtils.Delay(1000, Indicator.Text);
//  Sys.Desktop.KeyUp(0x09);
//  Sys.Desktop.KeyUp(0x09);
//  
//  var job = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
//  job.Click();
//  job.setText(jobNumber);
//  WorkspaceUtils.waitForObj(job);
//  WorkspaceUtils.waitForObj(table);
//
//var i=0;
//while((labels.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=60)){ 
//  aqUtils.Delay(100);
//  i++;
//  labels.Refresh();
//}
//if(labels.getText().OleValue.toString().trim().indexOf("results")==-1){ 
// ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
//}
//
//  var flag=false;
//  for(var v=0;v<table.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
//      flag=true;
//      break;
//    }
//    else{ 
//      table.Keys("[Down]");
//    }
//  }
//  
//  if(flag){
//  ReportUtils.logStep("INFO", "Job is listed in table to for Job Invoice Allocation");
//  ReportUtils.logStep_Screenshot("");
//  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy for Job Invoice Allocation"); 
//  closeFilter.Click();
//  aqUtils.Delay(1000, Indicator.Text);
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//    var clientApproved = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
//  WorkspaceUtils.waitForObj(clientApproved);
//  if(clientApproved.background!=10674625){
//    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
//    }
//    else{ 
//    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
//    }
//  var workingEstimate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
//  WorkspaceUtils.waitForObj(workingEstimate);
//  if(workingEstimate.background!=10674625){
//    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
//    }
//    else{ 
//    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
//    }
//  var lastInvoice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
//  var totalInvoice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
//  var billingPrice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
//  var netInvoiceOnAcc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
//  
//  var Budgeting = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
//  WorkspaceUtils.waitForObj(Budgeting);
//  Budgeting.Click();
//  aqUtils.Delay(100, Indicator.Text);
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//  var Estimate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
//  Estimate.Keys("Client Approved Estimate");
//  aqUtils.Delay(100, Indicator.Text);
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//  var FullBudget = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
//  WorkspaceUtils.waitForObj(FullBudget);
//  FullBudget.Click();
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//  var BudgetGrid = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//  WorkspaceUtils.waitForObj(BudgetGrid);
//  var ii=0;
//  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
//    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")||(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")){ 
////      if(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf("T")!=-1)
//         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim();
//         Log.Message(Estimatelines[ii]);
//         ii++;
////      else
////         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(7).OleValue.toString().trim();    
//    }
//  }
//  var Invoicing = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
//  WorkspaceUtils.waitForObj(Invoicing);
//  Invoicing.Click();
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//  var invoiceAllocation = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel;
////  var invoiceAllocation = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
//
//  for(var i=0;i<invoiceAllocation.ChildCount;i++){ 
//  if((invoiceAllocation.Child(i).isVisible())&&(invoiceAllocation.Child(i).text="Job Invoice Allocation")){
//    invoiceAllocation = invoiceAllocation.Child(i);
//    if(invoiceAllocation.JavaClassName=="TabControl"){ 
//    Log.Message(invoiceAllocation.FullName);
//    invoiceAllocation.Click();
//    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//    }else{ 
//    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//    }
//    
//    }
//    else{
//    Log.Message(invoiceAllocation.FullName);
//    invoiceAllocation.Click();
//    aqUtils.Delay(2000, Indicator.Text);;
////    invoiceAllocation.PopupMenu.Click("Job Invoice Allocation");
//    var popUp = Aliases.Maconomy.InvoiceAllocation_Popup.LightweightContainer;
//    Sys.HighlightObject(popUp);
//    ImageRepository.ImageSet.JobInvoiceAllocation.Click();
//    ReportUtils.logStep_Screenshot("");
//    aqUtils.Delay(100, "Job Invoice Allocation");
//    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//    }else{ 
//    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//    }
//    }
//    break;
//  }
//}
//var balance = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
//if(balance.getText()!="0.00"){
////  var StandardView = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
////  StandardView.Click();
//
//var NetBal = balance.getText().OleValue.toString().trim();
//NetBal = parseFloat(NetBal.replace(/,/g, ''));
//NetBal = NetBal.toFixed(2);
//
//  var tableGrid = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//  Sys.HighlightObject(tableGrid);
//  ImageRepository.ImageSet.Maximize1.Click();
//  Log.Message("Balance :"+NetBal)
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//  var Total = [];
//  if(NetBal.indexOf("-")!=-1){ 
//
////    for(var j=0;j<Estimatelines.length;j++){
////      var temp = 0;
////      var split_text = Estimatelines[j].split("*");
////    for(var i=0;i<tableGrid.getItemCount();i++){ 
////      if(EnvParams.Country.toUpperCase()=="INDIA"){
////      if(tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[0])!=-1){
////        var amt = parseFloat(tableGrid.getItem(i).getText_2(3).OleValue.toString().trim().replace(/,/g, ''));
////        amt = amt.toFixed(2);
////        temp= parseFloat(temp)+parseFloat(amt);
////      }
////      }else{ 
////      if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[0])!=-1){ 
////        
////      }
////      }
////    }
////    Total[j]= split_text[0]+"*"+temp;
////    }
////    for(var j=0;j<Total.length;j++){
////      Log.Message(Total[j]);
////      }
////    for(var j=0;j<Total.length;j++){
////      var totText = Total[j].split("*");
////      for(var k=0;k<Estimatelines.length;k++){
////        var split_text = Estimatelines[k].split("*");
////        if(totText[0].indexOf(split_text[0])!=-1){
////          totText[1] = totText[1].toFixed(2);
////          Estimatelines[4] = Estimatelines[4].toFixed(2);
////          totText[1] = parseFloat(totText[1])
////          Estimatelines[4] = parseFloat(Estimatelines[4])
////          if(Estimatelines[4]>totText[1]){ 
////                for(var i=0;i<tableGrid.getItemCount();i++){ 
////                  if(EnvParams.Country.toUpperCase()=="INDIA"){
////                    if(tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(totText[0])!=-1){
////                    var amt = parseFloat(tableGrid.getItem(i).getText_2(3).OleValue.toString().trim().replace(/,/g, ''));
////                    amt = amt.toFixed(2);
////                    amt= parseFloat(amt);
//////                      if()
////                      }else{ 
////                       tableGrid.Keys("[Down]") ;
////                      }
////                      }
////                      }
////          
////          
////          
////          
////          }
////        
////      }
////      }
////  }
//  
//  }
//else{ 
//      for(var j=0;j<Estimatelines.length;j++){
//      var temp = 0;
//      var split_text = Estimatelines[j].split("*");
//      for(var i=0;i<tableGrid.getItemCount();i++){ 
//      if(EnvParams.Country.toUpperCase()=="INDIA"){
//      if(tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[0])!=-1){
//      var amt = parseFloat(tableGrid.getItem(i).getText_2(3).OleValue.toString().trim().replace(/,/g, ''));
//      amt = amt.toFixed(2);
//      temp= parseFloat(temp)+parseFloat(amt);
//      }
//      }else{ 
//      if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[0])!=-1){ 
//      var amt = parseFloat(tableGrid.getItem(i).getText_2(2).OleValue.toString().trim().replace(/,/g, ''));
//      amt = amt.toFixed(2);
//      temp= parseFloat(temp)+parseFloat(amt);
//      }
//      }
//      }
//      Total[j]= split_text[0]+"*"+temp;
//      }
//      for(var j=0;j<Total.length;j++){
//      Log.Message(Total[j]);
//      }
//// Select Work Code
//      for(var i=0;i<tableGrid.getItemCount();i++){
//        
//      for(var j=0;j<Estimatelines.length;j++){
//      var temp = 0;
//      var split_text = Estimatelines[j].split("*");
//        if(EnvParams.Country.toUpperCase()=="INDIA"){
//          if((tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[0])!=-1) &&
//          (tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf("HSN Code")!=-1)){
//            
//            var Entries = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
//            WorkspaceUtils.waitForObj(Entries);
//            Entries.Click();
//            for(var k=0;k<Total.length;k++){
//              var temp = Total[k].split("*");
//            if(temp[0]==split_text[0]){ 
//              if(temp[1]=="0"){ 
//                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(add);
//                add.Click();
//                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//                WorkspaceUtils.waitForObj(JobNo);
//                JobNo.Click();
//                Sys.Desktop.KeyDown(0x09); // Press Ctrl
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyUp(0x09);
//                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(desc);
//                desc.Click();
//                desc.setText(split_text[1])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
//                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//                WorkspaceUtils.waitForObj(emp);
//                emp.Click();
//          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
//                if((EmpNo!="")&&(EmpNo!=null)){
//                emp.HoverMouse();
//                emp.Click();
//                WorkspaceUtils.SearchByValue(emp,"Employee",EmpNo,"Employee Number");
//                }
//                }
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(qty);
//                qty.Click();
//                qty.setText(split_text[2])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(unitprice);
//                unitprice.Click();
//                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//              }
//              
//            }else{ 
//                var amt = parseFloat(split_text[4]);
//                amt = parseFloat(amt.toFixed(2));
//                Log.Message("amt :"+amt)
//                var tot = parseFloat(temp[1]);
//                tot = parseFloat(tot.toFixed(2));
//                Log.Message("tot :"+tot)
//                var unit = parseFloat(split_text[3]);
//                unit = parseFloat(Qt.toFixed(3));
//                Log.Message("unit :"+unit)
//                // OverAll workCode Ammount " - " Po or Expernse Amount already in there in maconomy
//                var curentBal = amt - tot;
//                Log.Message("curentBal :"+curentBal)
//                var Quantity = curentBal / unit;
//                Log.Message("Quantity :"+Quantity)
//                if(Quantity.indexOf(".")!=-1){
//                var BeforeQty = parseFloat(Quantity.toString().split(".")[0]);
//                var AfterQty = parseFloat(num.toString().split(".")[1])
//// Amount Not 0 -> Adding Lines with Demical Quantity
//                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(add);
//                add.Click();
//                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//                WorkspaceUtils.waitForObj(JobNo);
//                JobNo.Click();
//                Sys.Desktop.KeyDown(0x09); // Press Ctrl
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyUp(0x09);
//                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(desc);
//                desc.Click();
//                desc.setText(split_text[1])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
//                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//                WorkspaceUtils.waitForObj(emp);
//                emp.Click();
//          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
//                if((EmpNo!="")&&(EmpNo!=null)){
//                emp.HoverMouse();
//                emp.Click();
//                WorkspaceUtils.SearchByValue(emp,"Employee",EmpNo,"Employee Number");
//                }
//                }
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(qty);
//                qty.Click();
//                qty.setText(BeforeQty)
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(unitprice);
//                unitprice.Click();
//                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
////Second Line
//                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(add);
//                add.Click();
//                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//                WorkspaceUtils.waitForObj(JobNo);
//                JobNo.Click();
//                Sys.Desktop.KeyDown(0x09); // Press Ctrl
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyUp(0x09);
//                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(desc);
//                desc.Click();
//                desc.setText(split_text[1])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
//                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//                WorkspaceUtils.waitForObj(emp);
//                emp.Click();
//          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
//                if((EmpNo!="")&&(EmpNo!=null)){
//                emp.HoverMouse();
//                emp.Click();
//                WorkspaceUtils.SearchByValue(emp,"Employee",EmpNo,"Employee Number");
//                }
//                }
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(qty);
//                qty.Click();
//                qty.setText(AfterQty)
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(unitprice);
//                unitprice.Click();
//                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                
//                }else{
////Absolute Quantity
//                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(add);
//                add.Click();
//                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//                WorkspaceUtils.waitForObj(JobNo);
//                JobNo.Click();
//                Sys.Desktop.KeyDown(0x09); // Press Ctrl
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyUp(0x09);
//                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(desc);
//                desc.Click();
//                desc.setText(split_text[1])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
//                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//                WorkspaceUtils.waitForObj(emp);
//                emp.Click();
//          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
//                if((EmpNo!="")&&(EmpNo!=null)){
//                emp.HoverMouse();
//                emp.Click();
//                WorkspaceUtils.SearchByValue(emp,"Employee",EmpNo,"Employee Number");
//                }
//                }
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(qty);
//                qty.Click();
//                qty.setText(Quantity)
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
//                WorkspaceUtils.waitForObj(unitprice);
//                unitprice.Click();
//                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                
//                
//                
//                
//                
//                }
//            }
//            }
//            
//          }
////          else{ 
////            tableGrid.Keys("[Down]");
////          }          
//   
//        else{ 
//          if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[0])!=-1){
//            
//          }
////          else{ 
////            tableGrid.Keys("[Down]");
////          }  
//        }
//      }
//      
//      }
//}
//
//
//  var Action = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.GroupToolItemControl;
//  aqUtils.Delay(2000, Indicator.Text);;
//  Action.PopupMenu.Click("Approve");
//  ReportUtils.logStep_Screenshot("");
//  aqUtils.Delay(100, "Approve is Clicked");
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//}
//  }
//  }
//
//}


function gotoAllocation(){ 
  var allJobs = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
allJobs.Click();


var labels = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
    labels = labels.Child(i);
    break;
  }
}

WorkspaceUtils.waitForObj(labels);

  var table = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
  
  var job = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);

var i=0;
while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==jobNumber){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to for Job Invoice Allocation");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+jobNumber+") is available in maconommy for Job Invoice Allocation"); 
  closeFilter.Click();
  
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var clientApproved = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(clientApproved);
  if(clientApproved.background!=10674625){
    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
    }
    
  var workingEstimate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
  WorkspaceUtils.waitForObj(workingEstimate);
  if(workingEstimate.background!=10674625){
    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
    }
    
  var lastInvoice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  var totalInvoice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  var billingPrice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
  var netInvoiceOnAcc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  
  var Budgeting = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Budgeting);
  Budgeting.Click();
  
  aqUtils.Delay(100, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  var Quote = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;
  WorkspaceUtils.waitForObj(Quote);
  Quote.Click();
  
//  var Estimate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
//  Estimate.Keys("Client Approved Estimate");
//  aqUtils.Delay(100, Indicator.Text);
//  
//  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//  }else{ 
//   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//  }
//  
//  var FullBudget = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
//  WorkspaceUtils.waitForObj(FullBudget);
//  FullBudget.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
//  var BudgetGrid = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  var BudgetGrid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(BudgetGrid);
  var ii=0;
  
  for(var i=0;i<BudgetGrid.getItemCount();i++){ 
    if((BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")||(BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()!="")){ 
//         Estimatelines[ii] = BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(9).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(10).OleValue.toString().trim();
         Estimatelines[ii] = "WorkCode"+"*"+BudgetGrid.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(1).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(2).OleValue.toString().trim()+"*"+BudgetGrid.getItem(i).getText_2(3).OleValue.toString().trim();
         Log.Message(Estimatelines[ii]);
         ii++;
    }
  }
  
//  var Invoicing = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
  var Invoicing = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite11.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8)
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var invoiceAllocation = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel;
  for(var i=0;i<invoiceAllocation.ChildCount;i++){ 
  if((invoiceAllocation.Child(i).isVisible())&&(invoiceAllocation.Child(i).text=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Invoice Allocation").OleValue.toString().trim())){
    invoiceAllocation = invoiceAllocation.Child(i);
    if(invoiceAllocation.JavaClassName=="TabControl"){ 
    Log.Message(invoiceAllocation.FullName);
    invoiceAllocation.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }else{ 
    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
    }
    
    }
    else{
    Log.Message(invoiceAllocation.FullName);
    invoiceAllocation.Click();
    aqUtils.Delay(2000, Indicator.Text);;
    var popUp = Aliases.Maconomy.InvoiceAllocation_Popup.LightweightContainer;
    Sys.HighlightObject(popUp);
    if(ImageRepository.ImageSet.JobInvoiceAllocation.Exists()){
    ImageRepository.ImageSet.JobInvoiceAllocation.Click();
    ReportUtils.logStep_Screenshot("");
    }
    if(ImageRepository.ImageSet.Allocation_Wip.Exists()){
    ImageRepository.ImageSet.Allocation_Wip.Click();
    ReportUtils.logStep_Screenshot("");
    }
    aqUtils.Delay(100, "Job Invoice Allocation");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }else{ 
    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
    }
    }
    break;
  }
}// invoiceAllocation For loop

var balance = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
if(balance.getText()!="0.00"){
  var NetBal = balance.getText().OleValue.toString().trim();
NetBal = parseFloat(NetBal.replace(/,/g, ''));
NetBal = NetBal.toFixed(2);

  var tableGrid = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(tableGrid);
  ImageRepository.ImageSet.Maximize1.Click();
  Log.Message("Balance :"+NetBal)
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var Total = [];
  if(NetBal.indexOf("-")!=-1){ 
    

  }
  else{ 
      for(var j=0;j<Estimatelines.length;j++){
      var temp = 0;
      var split_text = Estimatelines[j].split("*");
      for(var i=0;i<tableGrid.getItemCount();i++){ 
      if(EnvParams.Country.toUpperCase()=="INDIA"){
//      if(tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[0])!=-1){
      if(tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[1])!=-1){
      var amt = parseFloat(tableGrid.getItem(i).getText_2(3).OleValue.toString().trim().replace(/,/g, ''));
      amt = amt.toFixed(2);
      temp= parseFloat(temp)+parseFloat(amt);
      }
      }else{ 
//      if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[0])!=-1){ 
      if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[1])!=-1){ 
      var amt = parseFloat(tableGrid.getItem(i).getText_2(2).OleValue.toString().trim().replace(/,/g, ''));
      amt = amt.toFixed(2);
      temp= parseFloat(temp)+parseFloat(amt);
      }
      }
      }
//      Total[j]= split_text[0]+"*"+temp;
      Total[j]= split_text[1]+"*"+temp;
      }
      
      for(var j=0;j<Total.length;j++){
      Log.Message(Total[j]);
      }
var TableDes = "";
      for(var i=0;i<tableGrid.getItemCount();i++){
      for(var j=0;j<Estimatelines.length;j++){
      var temp = 0;
      var split_text = Estimatelines[j].split("*");
        if(EnvParams.Country.toUpperCase()=="INDIA"){
//          if((tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[0])!=-1) &&
//          (tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf("Default")==-1)){
            
          if((tableGrid.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(split_text[1])!=-1) &&
          (tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf("Default")==-1)){
            TableDes = tableGrid.getItem(i).getText_2(1).OleValue.toString().trim();
            Log.Message("TableDes :"+TableDes)
          
            // Add Entries
            
//            var Entries = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
//            WorkspaceUtils.waitForObj(Entries);
//            Entries.Click();
            for(var k=0;k<Total.length;k++){
              var temp = Total[k].split("*");
              var t1 = parseFloat(temp[1]);
              t1 = parseFloat(t1.toFixed(2));  //Allocated Amount
              split_text[4] = split_text[4].replace(/,/g, '');
              var t2 = parseFloat(split_text[4]);
              t2 = parseFloat(t2.toFixed(2));  //Estimate Amount
              Log.Message(temp[0])
              Log.Message(split_text[0])
              Log.Message(t1)
              Log.Message(t2)
//              Log.Message((temp[0]==split_text[0]));
              Log.Message((temp[0]==split_text[1]));
              Log.Message(t1!=t2);
//            if((temp[0]==split_text[0])&&(t1!=t2)){ 
            if((temp[0]==split_text[1])&&(t1!=t2)){ 
            var Entries = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
            WorkspaceUtils.waitForObj(Entries);
            Entries.Click();
            
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
            }else{ 
            ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
            }
            Log.Message(temp[1])
              if(temp[1]=="0"){ 
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(split_text[2])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
              ImageRepository.ImageSet.Close_Down.Click();
              allocateMainTable(tableGrid);
////mainTable
//              Sys.HighlightObject(tableGrid);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
//                WorkspaceUtils.waitForObj(allocate)
//                allocate.Keys("Allocate");
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//              Sys.HighlightObject(tableGrid);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
////Save Main Table
              }            
              else{
                Log.Message(Estimatelines[j]);
                Log.Message(split_text[0])
                Log.Message(split_text[1])
                Log.Message(split_text[2])
                Log.Message(split_text[3])
                Log.Message(split_text[4])
                split_text[4] = split_text[4].replace(/,/g, '');
                split_text[3] = split_text[3].replace(/,/g, '');
                var amt = parseFloat(split_text[4]);
                amt = parseFloat(amt.toFixed(2));
                Log.Message("amt :"+amt)
                var tot = parseFloat(temp[1]);
                tot = parseFloat(tot.toFixed(2));
                Log.Message("tot :"+tot)
                var unit = parseFloat(split_text[3]);
                unit = parseFloat(unit.toFixed(3));
                Log.Message("unit :"+unit)
                // OverAll workCode Ammount " - " Po or Expernse Amount already in there in maconomy
                var curentBal = amt - tot;
                Log.Message("curentBal :"+curentBal)
                var Quantity = curentBal / unit;
                Log.Message("Quantity :"+Quantity)
                if(Quantity.toString().indexOf(".")!=-1){
                  
                var BeforeQty = parseFloat(Quantity.toString().split(".")[0]);
                var AfterQty = parseFloat(Quantity.toString().split(".")[1]);
                var AfterAmount;
                var start = true;
// Amount Not 0 -> Adding Lines with Demical Quantity
                if(BeforeQty>=2){ 
                  BeforeQty = BeforeQty-1;
                  var BeforeAmount = BeforeQty*unit;
                  Log.Message("BeforeAmount :"+BeforeAmount);
                  AfterAmount = curentBal - BeforeAmount;
                }else{ 
                  start = false;
                  AfterAmount = curentBal;
                }
                if(start){
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(BeforeQty)
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                
                }
//Second Line
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
//                qty.setText("0."+AfterQty)
                qty.setText("1")
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
//                unitprice.setText(split_text[3])
                unitprice.setText(AfterAmount)
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                ImageRepository.ImageSet.Close_Down.Click();
                allocateMainTable(tableGrid);
                  }else{ 
//Absolute
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(Quantity)
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                ImageRepository.ImageSet.Close_Down.Click();
                allocateMainTable(tableGrid);
                  }
              }
              }
                
          }
          }
          }
// for Other Country
          else{ // for Other Country
//          if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[0])!=-1){
          if(tableGrid.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(split_text[1])!=-1){
          TableDes = tableGrid.getItem(i).getText_2(0).OleValue.toString().trim();
            // Add Entries
            
//            var Entries = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
//            WorkspaceUtils.waitForObj(Entries);
//            Entries.Click();
            for(var k=0;k<Total.length;k++){
              var temp = Total[k].split("*");
              var t1 = parseFloat(temp[1]);
              t1 = parseFloat(t1.toFixed(2));  //Allocated Amount
              split_text[4] = split_text[4].replace(/,/g, '');
              var t2 = parseFloat(split_text[4]);
              t2 = parseFloat(t2.toFixed(2));  //Estimate Amount
              Log.Message(temp[0])
              Log.Message(split_text[0])
              Log.Message(t1)
              Log.Message(t2)
//              Log.Message((temp[0]==split_text[0]));
              Log.Message((temp[0]==split_text[1]));
              Log.Message(t1!=t2);
//            if((temp[0]==split_text[0])&&(t1!=t2)){ 
            if((temp[0]==split_text[1])&&(t1!=t2)){ 
            var Entries = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
            WorkspaceUtils.waitForObj(Entries);
            Entries.Click();
            
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
            }else{ 
            ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
            }
            
              if(temp[1]=="0"){ 
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(split_text[2])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
              ImageRepository.ImageSet.Close_Down.Click();
              allocateMainTable(tableGrid);
////mainTable
//              Sys.HighlightObject(tableGrid);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
//                WorkspaceUtils.waitForObj(allocate)
//                allocate.Keys("Allocate");
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//              Sys.HighlightObject(tableGrid);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
////Save Main Table
              }            
              else{
                Log.Message(Estimatelines[j]);
                Log.Message(split_text[0])
                Log.Message(split_text[1])
                Log.Message(split_text[2])
                Log.Message(split_text[3])
                Log.Message(split_text[4])
                split_text[4] = split_text[4].replace(/,/g, '');
                split_text[3] = split_text[3].replace(/,/g, '');
                var amt = parseFloat(split_text[4]);
                amt = parseFloat(amt.toFixed(2));
                Log.Message("amt :"+amt)
                var tot = parseFloat(temp[1]);
                tot = parseFloat(tot.toFixed(2));
                Log.Message("tot :"+tot)
                var unit = parseFloat(split_text[3]);
                unit = parseFloat(unit.toFixed(3));
                Log.Message("unit :"+unit)
                // OverAll workCode Ammount " - " Po or Expernse Amount already in there in maconomy
                var curentBal = amt - tot;
                Log.Message("curentBal :"+curentBal)
                var Quantity = curentBal / unit;
                Log.Message("Quantity :"+Quantity)
                if(Quantity.toString().indexOf(".")!=-1){
                  
                var BeforeQty = parseFloat(Quantity.toString().split(".")[0]);
                var AfterQty = parseFloat(Quantity.toString().split(".")[1]);
                var AfterAmount;
                var start = true;
// Amount Not 0 -> Adding Lines with Demical Quantity
                if(BeforeQty>=2){ 
                  BeforeQty = BeforeQty-1;
                  var BeforeAmount = BeforeQty*unit;
                  Log.Message("BeforeAmount :"+BeforeAmount);
                  AfterAmount = curentBal - BeforeAmount;
                }else{ 
                  start = false;
                  AfterAmount = curentBal;
                }
                if(start){
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(BeforeQty)
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                
                }
//Second Line
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
//                qty.setText("0."+AfterQty)
                qty.setText("1")
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
//                unitprice.setText(split_text[3])
                unitprice.setText(AfterAmount)
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                ImageRepository.ImageSet.Close_Down.Click();
                allocateMainTable(tableGrid);
                  }else{ 
//Absolute
                var add = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var JobNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(JobNo);
                JobNo.Click();
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(desc);
                desc.Click();
                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
//                if(split_text[0].indexOf("T")!=-1){
                if((TableDes.indexOf("T")==0)||(TableDes.indexOf("BT")==0)){
                var emp = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
                WorkspaceUtils.waitForObj(emp);
                emp.Click();
          //      if((Employee.getText()=="")||(Employee.getText()==null)){ 
                if((EmpNo!="")&&(EmpNo!=null)){
                emp.HoverMouse();
                emp.Click();
                WorkspaceUtils.SearchByValue(emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
                }
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var qty = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(Quantity)
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(split_text[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate)
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Allocate").OleValue.toString().trim());
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                ImageRepository.ImageSet.Close_Down.Click();
                allocateMainTable(tableGrid);
                  }
              }
              }
                
          }
          }
//-------------------------------------------------------------------------          
          }
          
          }
          Sys.HighlightObject(tableGrid);
          Sys.HighlightObject(tableGrid);
          Sys.HighlightObject(tableGrid);
          tableGrid.Keys("[Down]");
          }
      
      
        }
        var closeBar = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
        WorkspaceUtils.waitForObj(closeBar);
        closeBar.Click();
      } 
//      else{ //if balance is " ) "

var check_Bal = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
if(check_Bal.getText()=="0.00"){ 
  ValidationUtils.verify(true,true,"Amount for Allocation is Balanced")
}else{ 
  ValidationUtils.verify(true,false,"Amount for Allocation is not Balanced")
}
  var Action = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.GroupToolItemControl;
  Sys.HighlightObject(Action);
  Action.Click();
  aqUtils.Delay(2000, Indicator.Text);;
  Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submit").OleValue.toString().trim());
  ReportUtils.logStep_Screenshot("");
  aqUtils.Delay(100, "Submit is Clicked");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var Action = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite2.GroupToolItemControl;
  Sys.HighlightObject(Action);
  Action.Click();
  WorkspaceUtils.waitForObj(Action);
  aqUtils.Delay(2000, Indicator.Text);;
  Action.PopupMenu.Click(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim());
  ReportUtils.logStep_Screenshot("");
  aqUtils.Delay(100, "Approve is Clicked");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
  ImageRepository.ImageSet.Close_Down.Click();
  
  LatestTran = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite2.McGroupWidget.Composite.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(LatestTran);
  Log.Message(LatestTran.getText());
  LatestTran = LatestTran.getText();
  var standardView = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
  WorkspaceUtils.waitForObj(standardView);
  standardView.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
//      }
            
}//Flag
}// Main


function gotoGeneralJournal(){ 
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Lookups").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GL Lookups").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Lookups from General Ledger Menu");
TextUtils.writeLog("Entering into GL Lookups from General Ledger Menu");
}

function GlLookups(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var journal = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl2;
  WorkspaceUtils.waitForObj(journal);
  journal.Click();
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var labels = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
  WorkspaceUtils.waitForObj(labels);
  for(var i=0;i<labels.ChildCount;i++){ 
    if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Now showing").OleValue.toString().trim())!=-1)){
      labels = labels.Child(i);
      break;
    }
  }

  WorkspaceUtils.waitForObj(labels);
  var JornalNo = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  WorkspaceUtils.waitForObj(JornalNo);
  JornalNo.Click();
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  
  var firstTrans = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  WorkspaceUtils.waitForObj(firstTrans);
  firstTrans.Click();
  firstTrans.setText(LatestTran);
  var closeFilter = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  var table = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  
  var i=0;
while((labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "results").OleValue.toString().trim())==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(9).OleValue.toString().trim()==LatestTran){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
  if(flag){
  ReportUtils.logStep("INFO", "Transaction Number is available in Journal");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Latest Transaction("+LatestTran+") is available in maconommy for Job Invoice Allocation"); 
  closeFilter.Click();
  
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  
//  var entriesGrid = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//  var date = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget
//  1707603881

  var JournalEntries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  WorkspaceUtils.waitForObj(JournalEntries);
  Log.Message(JournalEntries.FullName);
  JournalEntries.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var printJournal = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  WorkspaceUtils.waitForObj(printJournal);
  printJournal.Click();
  
//  var layout = Aliases.Maconomy.Shell3.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  var layout = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Journal").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2)
  WorkspaceUtils.waitForObj(layout);
  layout.Keys("Standard");
  aqUtils.Delay(10000, "Printing Journal");
  var printLayout = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Journal").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Print Journal").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(printLayout);
  printLayout.Click();
  
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
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
saveAs.Click();
aqUtils.Delay(2000, Indicator.Text);

Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Journal is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);

  
  
}
  
}

function allocateMainTable(tableGrid){ 
  //mainTable
//              Sys.HighlightObject(tableGrid);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                var allocate = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
//                WorkspaceUtils.waitForObj(allocate)
//                allocate.Keys("Allocate");
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//                var save = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl;
//                WorkspaceUtils.waitForObj(save);
//                save.Click();
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//              Sys.HighlightObject(tableGrid);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                Sys.Desktop.KeyDown(0x10);
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x10);
//                Sys.Desktop.KeyUp(0x09);
//                aqUtils.Delay(100, Indicator.Text);
//                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
//    
//                }else{ 
//                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
//                }
//Save Main Table

}