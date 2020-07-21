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
var sheetName = "InvoicingCarryForward";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
var STIME = "";
var jobNumber,EmpNo = "";
var Language = "";

//InvoicingCarryForward

//Main Function
function CarryForward(){ 
TextUtils.writeLog("Partical Invoicing CarryForward Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
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
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);
 
}

Language = "";
Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}

Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "InvoicingCarryForward";
STIME = "";
Approve_Level =[];
ApproveInfo = [];
jobNumber,EmpNo = "";
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Invoice from Budget started::"+STIME);

getDetails();
gotoMenu();
gotoInvoicing();
WorkspaceUtils.closeAllWorkspaces();
for(var i=level;i<ApproveInfo.length;i++){
level=i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApprove(temp[1],temp[2],i);
}

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
}

function getDetails(){ 
sheetName ="InvoicingCarryForward";  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed for Invoice from Budget");
  
  EmpNo = ReadExcelSheet("Timesheet Employee No",EnvParams.Opco,"Data Management");
  if((EmpNo=="")||(EmpNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  EmpNo = ExcelUtils.getColumnDatas("Employee Number",EnvParams.Opco)
  }
  if((EmpNo=="")||(EmpNo==null))
  ValidationUtils.verify(false,true,"Employee Number is needed for Invoice from Budget");
  
 
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
Client_Managt.ClickItem("|Jobs");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Jobs");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}


function gotoInvoicing(){
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  var allJobs = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
allJobs.Click();


var labels = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;
WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf("Now showing")!=-1)){
    labels = labels.Child(i);
    break;
  }
}

WorkspaceUtils.waitForObj(labels);

  var table = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
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
  
  var job = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(jobNumber);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);

var i=0;
while((labels.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf("results")==-1){ 
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
  ReportUtils.logStep("INFO", "Job is listed in table to for Invoice OnAccount");
  ReportUtils.logStep_Screenshot("");
  TextUtils.InvoicingCarryForward("Job("+jobNumber+") is available in maconommy for Invoice OnAccount"); 
  closeFilter.Click();
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
  var clientApproved = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
  WorkspaceUtils.waitForObj(clientApproved);
  if(clientApproved.background!=10674625){
    ValidationUtils.verify(true,false,"Client Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Client Approved Estimate is Fully Approved")  
    }
  var workingEstimate = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Composite2.McTextWidget;
  WorkspaceUtils.waitForObj(workingEstimate);
  if(workingEstimate.background!=10674625){
    ValidationUtils.verify(true,false,"Working Approved Estimate is not Fully Approved")
    }
    else{ 
    ValidationUtils.verify(true,true,"Working Approved Estimate is Fully Approved")  
    }
  var lastInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McTextWidget;
  var totalInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
  var billingPrice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
  var netInvoiceOnAcc = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
  

  var Invoicing = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
  WorkspaceUtils.waitForObj(Invoicing);
  Invoicing.Click();
  

//  Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7)
//                   Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite3.PTabFolder.TabFolderPanel.TabControl             
  var iselection = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Log.Message(iselection.FullName)
  WorkspaceUtils.waitForObj(iselection);
  ReportUtils.logStep_Screenshot("");
  iselection.Click();
  TextUtils.writeLog("Entering into Invoice Selection Tab");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }


  var SelectionBilling = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  WorkspaceUtils.waitForObj(SelectionBilling);
  aqUtils.Delay(100, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
var oneWorkCode = false;
var CarryForward_Line = [];
var x=0;
var CarryForward_stat = false
   for(var i=1;i<=10;i++){
ExcelUtils.setExcelName(workBook, sheetName, true);
var wCodeID = ExcelUtils.getColumnDatas("WorkCode_"+i,EnvParams.Opco)
var Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco)
var Qly = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco)
var CostBase = ExcelUtils.getColumnDatas("Cost_"+i,EnvParams.Opco)
var OutwardHSN = ExcelUtils.getColumnDatas("HSN_"+i,EnvParams.Opco)
var Acct = ExcelUtils.getColumnDatas("Action_"+i,EnvParams.Opco)
if(Acct == "CarryForward"){ 
  CarryForward_stat = true;
}
if((wCodeID!="")&&(wCodeID!=null)){
            oneWorkCode = true;
  if(((Desp!="")&&(Desp!=null))){ 
    if(((Qly!="")&&(Qly!=null))){ 
      if(((CostBase!="")&&(CostBase!=null))){ 
      CarryForward_Line[x] = wCodeID+"*"+Desp+"*"+Qly+"*"+CostBase+"*";
        if(((OutwardHSN!="")&&(OutwardHSN!=null))){ 
        CarryForward_Line[x] = wCodeID+"*"+Desp+"*"+Qly+"*"+CostBase+"*"+OutwardHSN+"*"+Acct+"*";
          }else{ 
            if(EnvParams.Country.toUpperCase()=="INDIA"){ 
             ValidationUtils.verify(true,false,"HSN_"+i+" Code is required to Write-Off") 
            }else{ 
             CarryForward_Line[x] = wCodeID+"*"+Desp+"*"+Qly+"*"+CostBase+"*"+Acct+"*"; 
            }
          }
          x++;
        }
        else{ 
            ValidationUtils.verify(true,false,"Cost_"+i+" is required to Write-Off") 
          }
      }
        else{ 
            ValidationUtils.verify(true,false,"Quantity_"+i+" is required to Write-Off") 
          }
    
  }
  else{ 
            ValidationUtils.verify(true,false,"Description_"+i+" is required to Write-Off") 
          }
  
}
}

if(!oneWorkCode){ 
  ValidationUtils.verify(true,false,"Atleast one workcode is required to write off");
}
if(!CarryForward_stat){ 
  ValidationUtils.verify(true,false,"Atleast one Write-Off Action is required");
}

  if(EnvParams.Country.toUpperCase()=="INDIA")
  Runner.CallMethod("IND_Partical_InvoicingCarryForward.CarryForward",CarryForward_Line,SelectionBilling);
  else{ 
    
  
   for(var i=0;i<CarryForward_Line.length;i++)
  Log.Message(CarryForward_Line[i])
  ImageRepository.ImageSet.Maximize1.Click();
    for(var t=0;t<SelectionBilling.getItemCount();t++){ 
      var match = true;
      for(var z=0;z<CarryForward_Line.length;z++){
        var temp = CarryForward_Line[z].split("*");
    if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])!=-1){ 
      if(match){
      var entries = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      }
  match = false;
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
            }else{ 
            ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
            }
            
                var add = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var EmpNo = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(EmpNo);
                EmpNo.Click();
                var EmployeeNum = Partial_InvoicingCarryForward.EmpNo;
                if(temp[0].indexOf("T")==0){
                if((EmployeeNum!="")&&(EmployeeNum!=null)){ 
                WorkspaceUtils.SearchByValue(EmpNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_InvoicingCarryForward.Language, "Employee").OleValue.toString().trim(),EmployeeNum,"Employee Number :"+EmployeeNum);
                }
                }
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
//                var workCode = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//                workCode.HoverMouse();
//                workCode.Click();
//                WorkspaceUtils.SearchByValue(workCode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_InvoicingCarryForward.Language, "Work Code").OleValue.toString().trim(),temp[0],"Work Code :"+temp[0]);
                
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
//                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
//                WorkspaceUtils.waitForObj(desc);
//                desc.Click();
//                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);

                var qty = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(temp[2])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(temp[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
Log.Message(temp[4])
Log.Message(temp[4]=="Write-Off");


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
                var allocate = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate);
                if(temp[4]=="Write-Off"){
                allocate.Keys("Write Off");
                }else{ 
                allocate.Keys("Invoice");
                }
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }


                }
                }
              ImageRepository.ImageSet.Close_Down.Click();
              Sys.HighlightObject(SelectionBilling);
      
      SelectionBilling.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
      if(temp[0].indexOf("T")==0){
      var Employee = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
      WorkspaceUtils.waitForObj(Employee);
      Employee.Click();
      if((EmployeeNum!="")&&(EmployeeNum!=null)){ 
      Employee.HoverMouse();
      Employee.Click();
      WorkspaceUtils.SearchByValue(Employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_InvoicingCarryForward.Language, "Employee").OleValue.toString().trim(),EmployeeNum,"Employee Number :"+EmployeeNum);
      }
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      
      var Save = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
      WorkspaceUtils.waitForObj(Save);
      for(var i=0;i<Save.ChildCount;i++){ 
        if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText=="Save Invoice Selection Line")){
          Save = Save.Child(i);
          ReportUtils.logStep_Screenshot("");
          Save.Click();
          break;
        }
      }
      }

      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      if(temp[0].indexOf("T")==0){
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      }
     
      SelectionBilling.Keys("[Down]");
    

  }
  
  var closeSelection = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
  WorkspaceUtils.waitForObj(closeSelection);
  closeSelection.Click(); 
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  }

  
TextUtils.writeLog("Write Off lines are Saved");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

aqUtils.Delay(100, "Approve is Clicked");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
aqUtils.Delay(100, "Billing Price");
var BudgetAmount = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
BudgetAmount = BudgetAmount.getText();
var Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.SingleToolItemControl;
WorkspaceUtils.waitForObj(Approve);
ReportUtils.logStep_Screenshot("");
Approve.Click();
TextUtils.writeLog("Approve is clicked");
aqUtils.Delay(1000, "Approve is Clicked");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

var DraftInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;

WorkspaceUtils.waitForObj(DraftInvoice);
ReportUtils.logStep_Screenshot("");
DraftInvoice.Click();
aqUtils.Delay(1000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

var draftNo = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(draftNo);
draftNo.Keys("[Tab][Tab][Tab]");
var billiablePrice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
WorkspaceUtils.waitForObj(billiablePrice);
billiablePrice.Click();
billiablePrice.setText(BudgetAmount);

var DraftTable = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var flag=false;
  for(var v=0;v<DraftTable.getItemCount();v++){ 
    if(DraftTable.getItem(v).getText_2(3).OleValue.toString().trim()==BudgetAmount){ 
      flag=true;
      break;
    }
    else{ 
      DraftTable.Keys("[Down]");
    }
  }
 ValidationUtils.verify(true,flag,"Invoice On Account is available to submit Draft")
  if(flag){
var CloseFilter = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
ReportUtils.logStep_Screenshot("");
CloseFilter.Click();
aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }

var SubmitDraft = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
//                  Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
//var SubmitDraft = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
  WorkspaceUtils.waitForObj(SubmitDraft);
  for(var i=0;i<SubmitDraft.ChildCount;i++){ 
    if((SubmitDraft.Child(i).isVisible())&&(SubmitDraft.Child(i).toolTipText=="Submit Draft")){
      WorkspaceUtils.waitForObj(SubmitDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      SubmitDraft.Child(i).Click();
      TextUtils.writeLog("Draft Invoice is submitted");
      break;
    }
  }
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }else{ 
   ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
  }
var PrintDraft;
     
  if(Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
  PrintDraft = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
  else
  PrintDraft = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;

              
  WorkspaceUtils.waitForObj(PrintDraft);
  for(var i=0;i<PrintDraft.ChildCount;i++){ 
    if((PrintDraft.Child(i).isVisible())&&(PrintDraft.Child(i).toolTipText=="Print Draft")){
      WorkspaceUtils.waitForObj(PrintDraft.Child(i));
      ReportUtils.logStep_Screenshot("");
      PrintDraft.Child(i).Click();
      break;
    }
  } 
  
TextUtils.writeLog("Print Draft is Clicked");
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);;
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "Print Invoice Editing"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Print Invoice Editing")!=-1){
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
ValidationUtils.verify(true,true,"Print Draft Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf")
    aqUtils.Delay(4000, Indicator.Text);
   

var appvBar = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
  
ImageRepository.ImageSet.Maximize.Click();
var DraftApproval = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
var ApproverTable = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
 var y=0;
for(var i=0;i<ApproverTable.getItemCount();i++){   
   var approvers="";
    if(ApproverTable.getItem(i).getText_2(6)!="Approved"){
    approvers = EnvParams.Opco+"*"+jobNumber+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
    Log.Message("Approver level :" +i+ ": " +approvers);
    Approve_Level[y] = approvers;
    y++;
    }
}
ReportUtils.logStep_Screenshot("");
var closeBar = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
Log.Message(OpCo2[2]);
Log.Message(Project_manager);
if(OpCo2[2]==Project_manager){
level = 1;

var Approve;
////var Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite;
//  if(Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.isVisible())
//  Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
// else
//  Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
            
 var ApproveStat = false;
Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText=="Approve Draft")){
    Approve = Approve.Child(i);
    ApproveStat =true;
    break;
  }
}

if(!ApproveStat){ 
Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText=="Approve Draft")){
    Approve = Approve.Child(i);
    break;
  }
} 
}
Log.Message(Approve.FullName)
WorkspaceUtils.waitForObj(Approve);
ReportUtils.logStep_Screenshot();
Approve.Click();
ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved Draft Invoice");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

var screen = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
//             Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
Log.Message(screen.FullName)
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-100);
//var ApvPerson = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
//var ApvPerson = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
var ApvPerson = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
Log.Message(ApvPerson.FullName)
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
while (((ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("APPROVED")==-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("YOU")==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

  if((ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("APPROVED")!=-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("YOU")!=-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1)){
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected")
  }
  
if(Approve_Level.length==1){
var appvBar = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
WorkspaceUtils.waitForObj(appvBar);
appvBar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
  
ImageRepository.ImageSet.Maximize.Click();
var DraftApproval = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();

var ApproverTable = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
ReportUtils.logStep_Screenshot();
var closeBar = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
//
var printStat = false;
var printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder;
  for(var i=0;i<printInvoice.ChildCount;i++){ 
    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).Name.indexOf("TabFolderPanel")!=-1)){
      printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
      printStat =true;
      break;
    }
  } 
  
  if(!printStat) 
  printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
  
//  if(Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.isVisible())
//  printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
// else
//  printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
                 
  WorkspaceUtils.waitForObj(printInvoice);
  for(var i=0;i<printInvoice.ChildCount;i++){ 
    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).toolTipText=="Print Invoice")){
      WorkspaceUtils.waitForObj(printInvoice.Child(i));
      ReportUtils.logStep_Screenshot("");
      printInvoice.Child(i).Click();
      break;
    }
  } 
  
    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
    aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*JobInvoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*JobInvoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("JobInvoice")!=-1){
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
ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(sFolder+SaveTitle+".pdf");
var textobj;
  try{
  var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj).OleValue.toString();
  textobj = textobj.substring(textobj.indexOf("Invoice No: ")+12);
  Log.Message("Invoice No:"+textobj.substring(0,textobj.indexOf("Invoice Date")))
  textobj = textobj.substring(0,textobj.indexOf("Invoice Date"));
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Write Off  Invoice No",EnvParams.Opco,"Data Management",textobj)
  TextUtils.writeLog("Write Off  Invoice No: "+textobj);



//}
}
}

}
//----------
  
  }
  
  }
  
  
function CredentialLogin(){ 
  var AppvLevl = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
    Log.Message(temp);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }
//WorkspaceUtils.closeAllWorkspaces();
}

function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
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
aqUtils.Delay(3000, "Waiting to Refresh ToDo's List");
if(ImageRepository.ImageSet.ToDos_Icon.Exists())
{ 
  
}else{ 
  ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}
var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Invoice Drafts (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Purchase Order from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Invoice Drafts Substitute) (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Invoice Drafts (Substitute) from To-Dos List");  
var listPass = false;   
  }
}  


if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Invoice Drafts by Type (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Invoice Drafts by Type from To-Dos List"); 
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Invoice Drafts by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Invoice Drafts by Type (Substitute) from To-Dos List"); 
var listPass = false;   
  }
} 
  }
  
}



function FinalApprove(JobNum,Apvr,lvl){ 

var table = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder;
WorkspaceUtils.waitForObj(table);
Sys.HighlightObject(table);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
if(Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(firstCell);
firstCell.setText(JobNum);
var closefilter = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;

var labels = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McPagingWidget;

WorkspaceUtils.waitForObj(labels);
for(var i=0;i<labels.ChildCount;i++){ 
  if((labels.Child(i).isVisible())&&(labels.Child(i).WndCaption.indexOf("Now showing")!=-1)){
    labels = labels.Child(i);
    break;
  }
}

WorkspaceUtils.waitForObj(labels);
var i=0;
while((labels.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  labels.Refresh();
}
if(labels.getText().OleValue.toString().trim().indexOf("results")==-1){ 
 ValidationUtils.verify(true,false,"Maconomy is loading continously......") 
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
WorkspaceUtils.waitForObj(table);
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==JobNum){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Draft Invoice is available in Approval List");
TextUtils.writeLog("Created Draft Invoice is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();

var Approve;

  if(Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.isVisible())
  Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite;
 else
  Approve = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
            
    
Sys.HighlightObject(Approve);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).toolTipText=="Approve Draft")){
    Approve = Approve.Child(i);
    break;
  }
}
//WorkspaceUtils.waitForObj(Approve);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();

ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
TextUtils.writeLog("Draft Invoice is Approved by "+Apvr);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}

              
var screen = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
WorkspaceUtils.waitForObj(screen);
  screen.Click();
  screen.MouseWheel(-100);

//  var Apv = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2
//Sys.HighlightObject(Apv);
//for(var i=0;i<Apv.ChildCount;i++){ 
//  if((Apv.Child(i).isVisible())&&(Apv.Child(i).Name.indexOf("McClumpSashForm")!=-1)){
//  Apv = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite
//    break;
//  }
//}
//  
//  var Apv = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite
//Sys.HighlightObject(Apv);
//for(var i=0;i<Apv.ChildCount;i++){ 
//  if((Apv.Child(i).isVisible())&&(Apv.Child(i).Name.indexOf("McClumpSashForm")!=-1)){
//  Apv = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite;
//    break;
//  }
//}

var Apv = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite;
Log.Message(Apv.FullName)
var ApvPerson;
for(var a=0;a<Apv.ChildCount;a++){ 
  if((Apv.Child(a).Visible)&&(Apv.Child(a).JavaClassName == "McTextWidget")){ 
    ApvPerson = Apv.Child(a);
    Log.Message("short");
    break;
  }
}
if((ApvPerson=="")||(ApvPerson==null)){ 
ApvPerson = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
Log.Message("Long")
}  
  
//                Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.ApproveStatus
//var ApvPerson = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget;
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;

while (((ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("APPROVED")==-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("YOU")==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))&&(i!=60))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();
}

if((ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("APPROVED")!=-1)||(ApvPerson.getText().OleValue.toString().trim().toUpperCase().indexOf("YOU")!=-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)!=-1)){
  ValidationUtils.verify(true,true,"Draft Invoice is Approved by :"+loginPer)
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer); 
  }else{ 
  TextUtils.writeLog("Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected"); 
  ValidationUtils.verify(true,false,"Draft Invoice is Approved by :"+loginPer+ "But its Not Reflected")
  }
  



if(Approve_Level.length==lvl+1){
  aqUtils.Delay(1000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
//var printInvoice = Aliases.Maconomy.InvoicePreparation.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite;
                  
var approvalBar = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
approvalBar.HoverMouse();
ReportUtils.logStep_Screenshot();
approvalBar.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
    ImageRepository.ImageSet.Maximize.Click();

var DraftApproval = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
WorkspaceUtils.waitForObj(DraftApproval);
DraftApproval.Click();
  

var ApproverTable = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
WorkspaceUtils.waitForObj(ApproverTable);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
ReportUtils.logStep_Screenshot();

var closeBar = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
WorkspaceUtils.waitForObj(closeBar);
closeBar.Click();
ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
}
//var printInvoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
//Sys.HighlightObject(printInvoice);
//printInvoice.Click();

var printStat = false;
var printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder;
  for(var i=0;i<printInvoice.ChildCount;i++){ 
    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).Name.indexOf("TabFolderPanel")!=-1)){
      printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
      printStat =true;
      break;
    }
  } 
  
  if(!printStat) 
  printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
  
//  if(Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2.isVisible())
//  printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.Composite2;
// else
//  printInvoice = Aliases.Maconomy.InvoicingCarryForward.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite2.PTabFolder.TabFolderPanel.Composite;
                 
  WorkspaceUtils.waitForObj(printInvoice);
  for(var i=0;i<printInvoice.ChildCount;i++){ 
    if((printInvoice.Child(i).isVisible())&&(printInvoice.Child(i).toolTipText=="Print Invoice")){
      WorkspaceUtils.waitForObj(printInvoice.Child(i));
      ReportUtils.logStep_Screenshot("");
      printInvoice.Child(i).Click();
      break;
    }
  } 
  
    TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
    aqUtils.Delay(5000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*"+".pdf - Adobe Acrobat Reader DC", 1).WndCaption.indexOf("Invoice")!=-1){
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
ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(sFolder+SaveTitle+".pdf");
var textobj;
  try{
  var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj).OleValue.toString();
  textobj = textobj.substring(textobj.indexOf("Invoice No: ")+12);
  Log.Message("Invoice No:"+textobj.substring(0,textobj.indexOf("Invoice Date")))
  textobj = textobj.substring(0,textobj.indexOf("Invoice Date"));
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Write Off Invoice No",EnvParams.Opco,"Data Management",textobj)
  TextUtils.writeLog("Write Off Invoice No: "+textobj);


}

  ValidationUtils.verify(true,true,"Draft Invoice is Approved by "+Apvr)
  
  
}
}

}