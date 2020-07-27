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
var sheetName = "Copy budget";
Indicator.Show();
Indicator.PushText("waiting for window to open");
var jobNumber = "";
var workcode,Internaldes,Quantity,CostBase,Billing,Employeecategory,Employee_Number,OutwardHSN,InwardHSN = "";
var Arrays = [];
var comapany="";
var ClientCurrency ="";
var count = true;
var Approve_Level = [];
var y=0;
var ApproveInfo = [];
var level =0;
var z="";

function getDetails(){
  Indicator.PushText("Reading Data from Excel");
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  Log.Message(jobNumber)
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  }  
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Modify Budget");
  Indicator.PushText("Playback");
}


 
  function goToJobs() {        
      var all_job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
      waitForObj(all_job);
      Sys.HighlightObject(all_job)
      all_job.Click();  
      var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
      var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
      var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
      SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
      SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
      SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
      SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
      SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
   
      companyFilter.forceFocus();
      companyFilter.setVisible(true);
      companyFilter.ClickM();
      table.Child(0).setText("^a[BS]");
      table.Child(0).setText(EnvParams.Opco);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09); // Press Ctrl
      aqUtils.Delay(1000, Indicator.Text);;
      Sys.Desktop.KeyDown(0x09); // Press Ctrl
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyUp(0x09);

      var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
      waitForObj(job);
      job.ClickM();
      table.Child(2).forceFocus();
      table.Child(2).setVisible(true);
      table.Child(2).setText("^a[BS]");
      table.Child(2).setText(jobNumber);
      aqUtils.Delay(1000, Indicator.Text);
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
      aqUtils.Delay(2000, Indicator.Text);
      Sys.HighlightObject(closeFilter);
      if(flag){
          closeFilter.Click();
          var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
          waitForObj(Budget);
           Sys.HighlightObject(Budget)
          Budget.Click();
        
        var show_budget = "";
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).isVisible())
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==1)
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText()=="Show Budget")
        show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).isVisible())
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").Index==1)    
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText()=="Show Budget"){
            show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
        }

        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible())
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").Index==1)    
        if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText()=="Show Budget"){
            show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
         }
        Sys.HighlightObject(show_budget);
        aqUtils.Delay(2000, Indicator.Text);
        show_budget.Keys("Client Approved Estimate"); 
        aqUtils.Delay(3000, Indicator.Text);
        ValidationUtils.verify(true,true,"Client Appproved Estimate is Selected");
        TextUtils.writeLog("Client Appproved Estimate is Selected");         
      }
      
 

 if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl2.isVisible()){
var copy =   Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl2;
Sys.HighlightObject(copy)
}
else{ 
var copy =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl4; 
Sys.HighlightObject(copy);
Sys.HighlightObject(copy)
}
 
 
 
 waitForObj(copy);
    Sys.HighlightObject(copy);
    copy.Click();
    
    var copypopup = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim());
      Sys.HighlightObject(copypopup);
      
    var copyfrom = Aliases.Maconomy.Shell3.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
    Sys.HighlightObject(copyfrom)  ;
    if(jobNumber!=""){
      copyfrom.Click();
      WorkspaceUtils.SearchByValues(copyfrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number");
//      WorkspaceUtils.SearchByValue(copyfrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
     }
  
  var copybutton = Aliases.Maconomy.Shell3.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy").OleValue.toString().trim());
    waitForObj(copybutton);
    Sys.HighlightObject(copybutton);
    copybutton.Click();
      
    }
 

function submit(){
  
// if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl3.isVisible()){
var Submit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl3;
Sys.HighlightObject(Submit)
//}
//else{ 
//var Submit =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl3
//Sys.HighlightObject(Submit);
//}
  
waitForObj(Submit);
var Add_Visible8 = true;
while(Add_Visible8){
aqUtils.Delay(2000, Indicator.Text);
if(Submit.isEnabled_2){
Add_Visible8 = false;
Sys.HighlightObject(Submit);
Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
Submit.Click();

ValidationUtils.verify(true,true,"Copy Budget is Submitted");

TextUtils.writeLog("Copy Budget is Submitted"); 
}
}
}


function approve(){
  
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl.isVisible()){
var approvepane = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl
}
else{ 
var approvepane =  Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
}
  
  waitForObj(approvepane)
  Sys.HighlightObject(approvepane);
  approvepane.Click();
  aqUtils.Delay(2000,Indicator.Text);
  ImageRepository.ImageSet0.Maximize.Click();
  
  if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2.isVisible()){
  var allaprove = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2;
    }
  else{ 
  var allaprove =  Aliases.Maconomy.CreateBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl
  }
  Sys.HighlightObject(allaprove);
  allaprove.Click();
  aqUtils.Delay(2000,Indicator.Text);
 
   if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl2.isVisible()){
  var Approval_table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
   }
  else{ 
  var Approval_table =  Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  }
  Sys.HighlightObject(Approval_table);
  Approval_table.Click();
  aqUtils.Delay(2000,Indicator.Text);
 
 Sys.HighlightObject(Approval_table);
  
   Sys.HighlightObject(Approval_table)
    for(var z=0;z<Approval_table.getItemCount();z++){ 
       approvers="";   
         if(Approval_table.getItem(z).getText_2(3)!="Approved"){
          approvers = EnvParams.Opco+"*"+jobNumber+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(5).OleValue.toString().trim();
            Log.Message("Approver level :" +z+ ": " +approvers);
//           if(Approval_table.getItem(z).getText_2(4).OleValue.toString().trim().indexOf("OpCo")!=-1){
//            var temp = (Approval_table.getItem(z).getText_2(4).OleValue.toString().trim()).split("-");
//                Log.Message(temp);
//                temp[0] = EnvParams.Opco;
//                approvers = temp[0]+" "+temp[1];
//                Log.Message(approvers)                           
//             }
//          Approve_Level[y] = EnvParams.Opco+"*"+jobNumber+"*"+approvers;
          Approve_Level[y] = approvers;
          Log.Message(Approve_Level[y])
          y++;
       }
    }
    
    TextUtils.writeLog("Finding approvers for Copy Budget");
  if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl.isVisible()){
  var closeCAList = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
  }
  else{ 
  var closeCAList =  Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
  }
      Sys.HighlightObject(closeCAList);
     closeCAList.HoverMouse();
     closeCAList.Click();
     ImageRepository.ImageSet0.Forward.Click();
    CredentialLogin();
    var OpCo2 = ApproveInfo[0].split("*");
    ExcelUtils.setExcelName(workBook, "Server Details", true);
    var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);

//    sheetName = "CreateClient";
    if(OpCo2[2]==Project_manager){
    level = 1;
     if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl4.isVisible()){
  var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite2.SingleToolItemControl4;
   }
  else{ 
  var Approve = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.SingleToolItemControl3;    
  }
  Sys.HighlightObject(Approve);
  Approve.Click();
  aqUtils.Delay(2000,Indicator.Text);
//    var Approve = 
    Sys.HighlightObject(Approve)
    if(Approve.isEnabled()){ 
    Approve.HoverMouse();
    ReportUtils.logStep_Screenshot();
    Approve.Click();
    aqUtils.Delay(8000, "Waiting for Approve");;
    ValidationUtils.verify(true,true,"Copy Budget is Approved by "+Project_manager)
    TextUtils.writeLog("Levels 0 has  Approved the Modify Budget");
    }
    } 
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
aqUtils.Delay(15000, Indicator.Text);
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
if(((temp.indexOf("Approve Job Budget by Type (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo por tipo (")!=-1))&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job Budget by Type from To-Dos List");
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(((temp.indexOf("Approve Job Budget by Type (Substitute) (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo por tipo (sustituto) (")!=-1))&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job Budget by Type (Substitute) from To-Dos List");
var listPass = true;   
  }
}  
if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(((temp.indexOf("Approve Job Budget (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo (")!=-1))&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Job Budget from To-Dos List");
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(((temp.indexOf("Approve Job Budget (Substitute) (")!=-1)||(temp.indexOf("Aprobar presupuesto de trabajo (sustituto) (")!=-1))&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Job Budget (Substitute) from To-Dos List");
var listPass = true;   
  }
} 
  }
//if(lvl==3){
//Client_Managt.ClickItem("|Approve Purchase Order (Substitute) (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Purchase Order (Substitute) (*)");
//TextUtils.writeLog("Entering into Approve Purchase Order (Substitute) from To-Dos List");
//}
//if(lvl==2){
//Client_Managt.ClickItem("|Approve Purchase Order (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Purchase Order (*)");
//TextUtils.writeLog("Entering into Approve Purchase Order from To-Dos List");
//}

}


function aprvBudget(ComId,JobNo,userNmae){ 
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet0.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet0.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);
} 

var alljobs = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
alljobs.Click();
 aqUtils.Delay(2000, Indicator.Text); 
var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
   var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
    SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
    SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).
    SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
    SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
    SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
   
    aqUtils.Delay(2000, Indicator.Text);    
    companyFilter.forceFocus();
    companyFilter.setVisible(true);
    companyFilter.ClickM();
    table.Child(0).setText("^a[BS]");
    table.Child(0).setText(ComId);
    aqUtils.Delay(2000, Indicator.Text);
    table.Child(0).Keys("[Tab][Tab]");

    var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    
    job.ClickM();
    table.Child(2).forceFocus();
    table.Child(2).setVisible(true);
    table.Child(2).setText("^a[BS]");
    table.Child(2).setText(JobNo);
    aqUtils.Delay(5000, Indicator.Text);
      var flag=false;
      for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(2).OleValue.toString().trim()==JobNo){ 
        flag=true;
        break;
      }
      else{ 
        table.Keys("[Down]");
      }
      }
    ValidationUtils.verify(true,true,"Modified Job is listed for Approval");
    
    if(table.getItemCount()>0){
//    Log.Message("Created Job is listed in table")
TextUtils.writeLog("Modified JobBudget is listed in Approval list");
ReportUtils.logStep_Screenshot("");
    closeFilter.Click();
    aqUtils.Delay(8000, Indicator.Text);
 var Budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
    Budget.Click();
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    ref.Refresh();
    aqUtils.Delay(5000, Indicator.Text);
 if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").Index==1)
    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    
    if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==1)
    var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
    Sys.HighlightObject(show_budget);   

    show_budget.Keys("Working Estimate");
    aqUtils.Delay(5000, Indicator.Text);
//    Delay(7000);
    var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 9);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
  Approve.HoverMouse();
ReportUtils.logStep_Screenshot("");
  Approve.Click();
ValidationUtils.verify(true,true,"Modified JobBudget is Approved by :"+userNmae)
//  TextUtils.writeLog("Modified JobBudget is Approved by :"+userNmae); 
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
}
    }

} 
   
function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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
    ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu"); 
  }
  

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}

function excel(){ 
  
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){   
      if(xlDriver.Value(colsList[idx])!=null){
            temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else
      {
        temp=temp+"*";
      } 
      }
     Arrays[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}




function CopyBudget(){ 
  TextUtils.writeLog("Copy Job Budget is Started"); 
Indicator.PushText("waiting for window to open");
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}
    Indicator.PushText("waiing for window to open");
    Language = EnvParams.LanChange(EnvParams.Language);
    WorkspaceUtils.Language = Language;
    STIME = WorkspaceUtils.StartTime();
    TextUtils.writeLog("Execution Start Time :"+STIME);
  try{  
    getDetails();
    goToJobMenuItem();
    goToJobs();
//    Multiplelines();
//    Copy();
    submit();
    approve();
//    closeAllWorkspaces();
      for(var i=level;i<ApproveInfo.length;i++){
      level = i;
      WorkspaceUtils.closeMaconomy();
      aqUtils.Delay(10000, Indicator.Text);
      var temp = ApproveInfo[i].split("*");
      Restart.login(temp[2]);
      aqUtils.Delay(5000, Indicator.Text);
      todo(temp[3]);
      aprvBudget(temp[0],temp[1],temp[2]);
      }
  }
    catch(err){
      Log.Message(err);
    }
  WorkspaceUtils.closeAllWorkspaces();
  }



function Clientgrid(){ 
var mainRoot = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");      
    var workcdee;
    var linestatus = false;
    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("3")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcdee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}
    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("2")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcdee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}


    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("4")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcdee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}

    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("5")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcdee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}
    if(!linestatus){
  for(var kk=0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
 var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("7")!=-1){  
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
for(var r =0;r<cc;r++){ 
var c = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5)
if((c.Child(r).Name.indexOf("McClumpSashForm")) && (c.Child(r).isVisible())){ 
workcdee = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
linestatus = true;
break;
}
}
if(linestatus){ 
  break;
}
}
}
}
}
}
return workcdee;
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

  
   
function Multiplelines(){ 
   var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
    var line = false;
  if(!line)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).ChildCount>=2)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible()){
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  line = true;
  }
  if(!line)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).ChildCount>=2)    
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).isVisible()){
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  line = true;
  }
  if(!line)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible()){
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  line = true;
  }
  if(!line)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible()){
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  line = true;
  }
   if(!line)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).isVisible()){
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  line = true;
  }
  
  Sys.HighlightObject(FullBudget);
  waitForObj(FullBudget);
  FullBudget.Click();  

}


function Copy(){
         if(!line)
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).isVisible()){
  var copy = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 12);    
 Sys.HighlightObject(copy);
   line = true;
  }
  
   if(!line)
 if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).isVisible()){
  var copy = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 12);    
   Sys.HighlightObject(copy);
  line = true;
  } 

  
     if(!line)
 if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).isVisible()){
  var copy = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 12);    
  Sys.HighlightObject(copy);
  line = true;
  }
     waitForObj(copy);
    Sys.HighlightObject(copy);
    copy.Click();
    
    var copypopup = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy Budget").OleValue.toString().trim());
      Sys.HighlightObject(copypopup);
      
    var copyfrom = Aliases.Maconomy.Shell3.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
    Sys.HighlightObject(copyfrom)  ;
    if(jobNumber!=""){
      copyfrom.Click();
      WorkspaceUtils.SearchByValues(copyfrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),jobNumber,"Job Number");
//      WorkspaceUtils.SearchByValue(copyfrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
     }
  
  var copybutton = Aliases.Maconomy.Shell3.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Copy").OleValue.toString().trim());
    waitForObj(copybutton);
    Sys.HighlightObject(copybutton);
    copybutton.Click();
}



function LogReport_name(ExcelData,value,JG){ 
var compStatus = "";
      for(var exl =0;exl<ExcelData.length;exl++){
        var splits = []; 
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
      if(splits[0]==value.toString().trim()){ 
        compStatus = ExcelData[exl]+"_"+JG;
        break;
      }
      }
return compStatus
}