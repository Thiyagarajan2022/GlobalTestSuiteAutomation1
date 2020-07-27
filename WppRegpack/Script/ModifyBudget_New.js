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
var sheetName = "Modify budget";
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
        show_budget.Keys("Working Estimate"); 
        ValidationUtils.verify(true,true,"Working Estimate is Selected");
        TextUtils.writeLog("Working Estimate is Selected");         
      }
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
    

      var addlines = false;
     if(!addlines)
  if((Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid).isVisible())
  {
  var grid = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;   
  addlines = true;
   }
       if(!addlines)
  if((NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid).isVisible())
  {
  var grid =NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  addlines = true;
   }
    grid.HoverMouse();
     Sys.HighlightObject(grid);
           
            var modifygridrow = grid.getItemCount()
            var modifygridcolumn = grid.getColumnCount()
            
    
 ///------Comparing each lines----   
    
      for( var i=0; i<modifygridrow-1; i++){
      for(z=1;z<=10;z++){
            ExcelUtils.setExcelName(workBook, sheetName, true);
            workcode = ExcelUtils.getColumnDatas("WorkCode_"+z,EnvParams.Opco)
            Log.Message(workcode)
            Internaldes = ExcelUtils.getColumnDatas("Description_"+z,EnvParams.Opco)
            Quantity = ExcelUtils.getColumnDatas("Quantity_"+z,EnvParams.Opco)
            CostBase = ExcelUtils.getColumnDatas("Cost_"+z,EnvParams.Opco)
//            Billing = ExcelUtils.getColumnDatas("Billing_"+z,EnvParams.Opco)
            Employeecategory = ExcelUtils.getColumnDatas("Employeecategory_"+z,EnvParams.Opco)
            Employee_Number = ExcelUtils.getColumnDatas("Employee_Number_"+z,EnvParams.Opco)
            OutwardHSN = ExcelUtils.getColumnDatas("Outward HSN_"+z,EnvParams.Opco)
            InwardHSN = ExcelUtils.getColumnDatas("Inward HSN_"+z,EnvParams.Opco)
            
            
            var workcdee = grid.getItem(i).getText_2(0);
            
          if(grid.getItem(i).getText_2(0)==workcode){                              
              if (i!=0){
                 Sys.Desktop.KeyDown(0x28);
                 Sys.Desktop.KeyUp(0x28);
              }                              
                 
                var Interdes = grid.getItem(i).getText_2(1);
                if(grid.getItem(i).getText_2(1)==Internaldes){
                  ValidationUtils.verify(true,true,"Given Internal Description in datasheet is as same as Value in Maconomy")
                }
                else{
                  renew()
                  break
                }
                  
                var empcat = grid.getItem(i).getText_2(4);
                if(grid.getItem(i).getText_2(4)==Employeecategory){
                  ValidationUtils.verify(true,true,"Given Employee Category in datasheet is as same as Value in Maconomy");
                }else{
                  renew()
                  break 
                }                 
                var empnum = grid.getItem(i).getText_2(5);
                if(grid.getItem(i).getText_2(5)==Employee_Number){
                  ValidationUtils.verify(true,true,"Given Employee Number in datasheet is as same as Value in Maconomy");
                } 
                else{
                  renew()
                  break
                } 
          
                var ActualQty = grid.getItem(i).getText_2(6).OleValue.toString().trim(); 
                ActualQty =  Math.trunc(ActualQty)
                if (ActualQty==Quantity){
                  ValidationUtils.verify(true,true,"Given Quantity in datasheet is as same as Value in Maconomy");
                } 
                else{
                  renew()
                  break
                } 
                
                var ActualCost = grid.getItem(i).getText_2(7).OleValue.toString().trim();  
                var temp = parseFloat(ActualCost.replace(/,/g, ''));
                ActualCost = Math.trunc(temp);          
                if (ActualCost==CostBase){
                  ValidationUtils.verify(true,true,"Given CostBase in datasheet is as same as Value in Maconomy");
                } 
                else{                  
                  renew()   
                  break         
                }   
                var billingval = grid.getItem(i).getText_2(9).OleValue.toString().trim();  
                var temp1 = parseFloat(billingval.replace(/,/g, ''));
                billingval = temp1;      
                if (billingval==CostBase){
                  ValidationUtils.verify(true,true,"Given Billing Price in datasheet is as same as Value in Maconomy");
                } 
                else{
                  renew()   
                  break         
                }                                                     
            } 
            
         }
        }        
    
}
           
 //// ------- Delete and Adding the Updated the line------   
    function renew(){                  
    ReportUtils.logStep_Screenshot();
    var reopen = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.SingleToolItemControl;
    Sys.HighlightObject(reopen);
    waitForObj(reopen);
    reopen.Click();  
// ///----Delete button    
    ReportUtils.logStep_Screenshot();
    var delet = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    delet.HoverMouse(); 
    Sys.HighlightObject(delet);
     delet.HoverMouse();
     waitForObj(delet);
     delet.Click();           
     TextUtils.writeLog("Job Budget Line is Deleted"); 
     aqUtils.Delay(2000, "Job Budget Line is Deleted");
     var popup = Sys.Process("Maconomy").SWTObject("Shell", "Delete");
     Sys.HighlightObject(popup);
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Delete").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
     waitForObj(OK);
     OK.HoverMouse();
      OK.Click();     
     Log.Message("Job Line has Deleted");
    ReportUtils.logStep_Screenshot();
 /// ---- Adding the deleted line    
     

  var AddBudget = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
   waitForObj(AddBudget);
   AddBudget.HoverMouse();
  Sys.HighlightObject(AddBudget)
  AddBudget.HoverMouse();
  AddBudget.Click();
  
     Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    
//-----Work Code Selection---------    
var workcdee;

linestatus = false;
workcdee = Clientgrid().SWTObject("McValuePickerWidget", "");
if(workcode!=""){
addedlines = true;
  workcdee.Click();
  WorkspaceUtils.SearchByValue(workcdee,"Work Code",workcode,"Work Code :"+workcode);
         }else{ 
    ValidationUtils.verify(false,true,"WorkCode Needed to Modify JobBudget");
  }
  
aqUtils.Delay(2000, Indicator.Text);  
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Play back");
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Play back");
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    ReportUtils.logStep_Screenshot();
    
//-----Internal Description---------
linestatus = false;
var External_Description;
External_Description = Clientgrid().SWTObject("McTextWidget", "", 4);

    Sys.HighlightObject(External_Description);
    External_Description.Click();
    if(Internaldes!=""){
    External_Description.setText(Internaldes);
    ValidationUtils.verify(true,true,"External Description is Entered");
    }    
   
    aqUtils.Delay(1000, "Play back");
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, "Play back");
    ReportUtils.logStep_Screenshot();
//-----Employee Categories if required---------
//var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
var EmpCat;
linestatus = false;
EmpCat = Clientgrid().SWTObject("McValuePickerWidget", "");
  if(Employeecategory!=""){
      EmpCat.Click();
      WorkspaceUtils.SearchByValue(EmpCat,"Employee Category",Employeecategory,"Employee Category");
   }         
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
ReportUtils.logStep_Screenshot();
var empno;
linestatus = false;
empno = Clientgrid().SWTObject("McValuePickerWidget", "");

    Sys.HighlightObject(empno);
    if(Employee_Number!=""){
  empno.Click();
  WorkspaceUtils.SearchByValue(empno,"Employee",Employee_Number,"Employee Number");
         }         
    aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);
    ReportUtils.logStep_Screenshot();
//-----Quantity---------
//var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
var Quality;
linestatus = false;
 Quality = Clientgrid().SWTObject("McTextWidget", "", 2);

    Sys.HighlightObject(Quality);
    Quality.Click();
    if(Quantity!=""){
    Quality.setText(Quantity);
    ValidationUtils.verify(true,true,"Quality is Entered");
    }
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(2000, Indicator.Text);
    ReportUtils.logStep_Screenshot();
    
//-----Cost Base Only for Amount---------
    if(workcode.indexOf("T")==-1){
//var cc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).ChildCount;
var Cost_base;
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
linestatus = false;
Cost_base = Clientgrid().SWTObject("McTextWidget", "", 2);

    Sys.HighlightObject(Cost_base);
    Cost_base.Click();
    if(CostBase!=""){    
    Cost_base.setText(CostBase);
    ValidationUtils.verify(true,true,"Cost is Entered");
    }
    }
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    ReportUtils.logStep_Screenshot();
    
    if(workcode.indexOf("T")>-1){
var Billing_Price;
 Billing_Price = Clientgrid().SWTObject("McTextWidget", "", 2);
linestatus = false;

    Sys.HighlightObject(Billing_Price);
    Billing_Price.Click();
    if(CostBase!=""){      
    Billing_Price.setText(CostBase);
    ValidationUtils.verify(true,true,"Cost is Entered");
    }
    }
var Country = EnvParams.Country;
if(Country.toUpperCase()=="INDIA")
{
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    
var Outward_HSN;
Outward_HSN = Clientgrid().SWTObject("McValuePickerWidget", "");
linestatus = false;
    Sys.HighlightObject(Outward_HSN);
    if(OutwardHSN!=""){
  Outward_HSN.Click();
  WorkspaceUtils.SearchByValue(Outward_HSN,"Local Specification 8",OutwardHSN,"Outward HSN");
         }else{ 
    ValidationUtils.verify(false,true,"Outward_HSN Needed to create JobBudget");
  }

    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    
var Invard_HSN;
linestatus = false;
Invard_HSN = Clientgrid().SWTObject("McValuePickerWidget", "");
    Sys.HighlightObject(Invard_HSN);
    if(InwardHSN!=""){
    Invard_HSN.Click();
    WorkspaceUtils.SearchByValue(Invard_HSN,"Local Specification 9",InwardHSN,"Inward HSN");
         }else{ 
    ValidationUtils.verify(false,true,"Inward HSN Needed to create JobBudget");
  }  
}  
   
var Save = "";
var kk= 0;
var mainRoot = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
var line = false;
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("2")!=-1){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2)
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("3")!=-1){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }

if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("4")!=-1){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("5")!=-1){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }
  
if(!line){
  for(var kk= 0;kk<mainRoot.ChildCount;kk++){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible()){
  var tempName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).Name;
  if(tempName.indexOf("7")!=-1){
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).isVisible())
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).ChildCount>=2)  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).isVisible()){
  Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
  line = true;
  break;
  }
  }
  }
  }
  }
    
    Sys.HighlightObject(Save);
    Save.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    Save.Click();
    aqUtils.Delay(3000, Indicator.Text);
   
     
  ExcelUtils.setExcelName(workBook,"JobBudgetCreation", true);
 
//  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"WorkCode_"+z,"JobBudgetCreation",workcode);  
//  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Description_"+z,"JobBudgetCreation",Internaldes);
////  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Employeecategory_"+z,"JobBudgetCreation",Employeecategory);
////  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Employee_Number_"+z,"JobBudgetCreation",Employee_Number);
//  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Quantity_"+z,"JobBudgetCreation",Quantity);
//  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Cost_"+z,"JobBudgetCreation",CostBase);
//  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Outward HSN_"+z,"JobBudgetCreation",OutwardHSN);
//  ExcelUtils.WriteBudgetExcelSheet(EnvParams.Opco,"Inward HSN_"+z,"JobBudgetCreation",InwardHSN);
//  
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"WorkCode_"+z,"JobBudgetCreation",workcode);  
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Description_"+z,"JobBudgetCreation",Internaldes);
//  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Employeecategory_"+z,"JobBudgetCreation",Employeecategory);
//  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Employee_Number_"+z,"JobBudgetCreation",Employee_Number);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Quantity_"+z,"JobBudgetCreation",Quantity);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Cost_"+z,"JobBudgetCreation",CostBase);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Outward HSN_"+z,"JobBudgetCreation",OutwardHSN);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Inward HSN_"+z,"JobBudgetCreation",InwardHSN);
  
  
  var Revisionnum = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Revision.getText().OleValue.toString().trim();
  Log.Message(Revisionnum)
  ExcelUtils.setExcelName(workBook,"CreateQuote", true);
  ExcelUtils.WriteExcelSheet(EnvParams.Opco,"Revision","CreateQuote",Revisionnum);
  
   TextUtils.writeLog("Budget Lines are modified and saved"); 
   ValidationUtils.verify(true,true,"Job Budget Line has modified");               
} 


function submit(){
  
var Submit = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite2.Submit;

var Add_Visible8 = true;
while(Add_Visible8){
aqUtils.Delay(2000, Indicator.Text);
if(Submit.isEnabled_2){
Add_Visible8 = false;
Sys.HighlightObject(Submit);
Submit.HoverMouse();
ReportUtils.logStep_Screenshot("");
Submit.Click();

ValidationUtils.verify(true,true,"Modify Budget is Submitted");

TextUtils.writeLog("Modified Working Estimate is Submitted"); 
}
}
}


function approve(){
  var approvepane = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  approvepane.Click();
  aqUtils.Delay(2000,Indicator.Text);
  ImageRepository.ImageSet0.Maximize.Click();
  
  var allaprove = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.TabControl3;
  allaprove.Click();
  
  var Approval_table = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
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
    
    TextUtils.writeLog("Finding approvers for Modify Budget");
    var closeCAList = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
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
    var Approve = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite.SingleToolItemControl3;    
    Sys.HighlightObject(Approve)
    if(Approve.isEnabled()){ 
    Approve.HoverMouse();
    ReportUtils.logStep_Screenshot();
    Approve.Click();
    aqUtils.Delay(8000, "Waiting for Approve");;
    ValidationUtils.verify(true,true,"Modify Budget is Approved by "+Project_manager)
    TextUtils.writeLog("Levels 0 has  Approved the Modify Budget");
    }
    } 
}


function todo1(lvl){
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
    var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Entering into To-Dos List");
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
  
  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();
  
  aqUtils.Delay(15000, Indicator.Text);
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");

if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Job Budget by Type (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into "+temp+" from To-Dos List");   
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Job Budget by Type (Substitute) (")!=-1)&&(temp1.length==3)){
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+temp);    
TextUtils.writeLog("Entering into "+temp+" from To-Dos List"); 
  }
}
break;
}
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




function ModifyBudget(){ 
  TextUtils.writeLog("Modify Job Budget is Started"); 
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
    Multiplelines();
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