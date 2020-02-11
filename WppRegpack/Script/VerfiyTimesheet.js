//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Verify Timesheet";
var Language = "";
  Indicator.Show();
  Indicator.PushText("waiing for window to open");
ExcelUtils.setExcelName(workBook, sheetName, true);

var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "Verify Timesheet", true);


      var EmpNumber ="";
      var weekno="";
      var mon,tue,wed,thu,fri = "";
      var Jobnumber ="";
      var Jobname ="";


function verifytimesheet() {  
     
      Language = "";
      Language = EnvParams.Language;
        if((Language==null)||(Language=="")){
          ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
        }      
      Language = EnvParams.LanChange(Language);
      WorkspaceUtils.Language = Language;
      Log.Message(Language)       
      getDetails();      
      goToJobMenuItem();
      gotoTimesheetlookup(); 
      WorkspaceUtils.closeAllWorkspaces();
}





function getDetails(){
  
  ExcelUtils.setExcelName(workBook, sheetName, true);
  EmpNumber = ExcelUtils.getRowDatas("Employee Name",EnvParams.Opco)  
  if((EmpNumber=="")||(EmpNumber==null))
  ValidationUtils.verify(false,true,"Employee Name is needed to Verify Timesheet");
  
  ExcelUtils.setExcelName(workBook, sheetName, true);
  weekno = ExcelUtils.getRowDatas("Weekno",EnvParams.Opco) 
  if((weekno=="")||(weekno==null))
  ValidationUtils.verify(false,true,"Week No is needed to Verify Timesheet");
  
  Jobnumber = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco) 
  if((Jobnumber=="")||(Jobnumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Verify Timesheet");
  
//  mon = ExcelUtils.getRowDatas("Mon",EnvParams.Opco)
//  if((mon==null)||(mon=="")){ 
//  ValidationUtils.verify(false,true,"Time for Monday is Needed to Verify Timesheet");
//  }
//  tue = ExcelUtils.getRowDatas("Tue",EnvParams.Opco)
//  if((tue==null)||(tue=="")){ 
//  ValidationUtils.verify(false,true,"Time for Tuesday is Needed to Verify Timesheet");
//  }
//  wed = ExcelUtils.getRowDatas("Wed",EnvParams.Opco)
//  if((wed==null)||(wed=="")){ 
//  ValidationUtils.verify(false,true,"Time for Wednessday is Needed to Verify Timesheet");
//  }
//  thu = ExcelUtils.getRowDatas("Thu",EnvParams.Opco)
//  if((thu==null)||(thu=="")){ 
//  ValidationUtils.verify(false,true,"Time for Thursday is Needed to Verify Timesheet");
//  }
//  fri= ExcelUtils.getRowDatas("Fri",EnvParams.Opco)
//  if((fri==null)||(fri=="")){ 
//  ValidationUtils.verify(false,true,"Time for Friday is Needed to Verify Timesheet");
//  }
} 

function goToJobMenuItem(){
       ReportUtils.logStep_Screenshot("");
 TextUtils.writeLog("Verify Timwsheet Started"); 
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
    menuBar.DblClick();
     if(ImageRepository.ImageSet2.TE.Exists()){
       ImageRepository.ImageSet2.TE.Click();// GL
      }
     else if(ImageRepository.ImageSet2.TE1.Exists()){
       ImageRepository.ImageSet2.TE1.Click();
      }
     else{
       ImageRepository.ImageSet2.TE2.Click();
    }
    
      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
      var Modify_Budget;
       for(var i=1;i<=childCC;i++){ 
          Time_Expenses = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i);
          if(Time_Expenses.isVisible()){ 
          Time_Expenses = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", ""); 
          Time_Expenses.ClickItem("|Time Sheet Lookups");
          ReportUtils.logStep_Screenshot();
          Time_Expenses.DblClickItem("|Time Sheet Lookups");
       }
      }
     aqUtils.Delay(6000, Indicator.Text);
     
     ReportUtils.logStep("INFO", "Moved to Time Sheet Lookups from Time & Expenses Menu");
}



function gotoTimesheetlookup(){ 
     
      ReportUtils.logStep_Screenshot("Verify the Timesheet");
 TextUtils.writeLog("Verify Timesheet"); 
 
 
      ReportUtils.logStep("INFO", "Verify Time Sheet");
      aqUtils.Delay(2000, Indicator.Text);
      var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.table;
      Sys.HighlightObject(table);
      var employee = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.table.empno;
      employee.Click();
      employee.setText(EmpNumber);
      aqUtils.Delay(1000, Indicator.Text);
      employee.Keys("[Tab][Tab][Tab][Tab]");
      aqUtils.Delay(1000, Indicator.Text);
      var Weekno = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.table.weekno;
       Sys.HighlightObject(Weekno);
       Weekno.Click();
       Weekno.setText(weekno);
       aqUtils.Delay(1000, Indicator.Text);
        ReportUtils.logStep_Screenshot(); 
       
       var flag=false;
      for(var v=0;v<table.getItemCount();v++){ 
        if(table.getItem(v).getText_2(4).OleValue.toString().trim()==weekno){ 
          flag=true;                            
          break;
        }
        else{ 
          table.Keys("[Down]");
        }        
      }                         
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
         
      ValidationUtils.verify(true,true,"Approved Employee is available in the Maconomy");
      
      var line = false;
      if(!line){
        if(Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.isVisible())  
        {
        var timesheetline = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.timesheetline;
        line = true;
        }
      }
      
       if(!line){  
       if(Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.isVisible())
        {
        var timesheetline = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid;  
        line = true;
        }  
      }
     Sys.HighlightObject(timesheetline);       
    ReportUtils.logStep_Screenshot("");
      
      for(var i=0;i<timesheetline.getItemCount();i++){
        
        var jobno = timesheetline.getItem(i).getText_2(0).OleValue.toString().trim();        
          if(jobno==Jobnumber){
            ValidationUtils.verify(true,true,"Job Number is Matched");
          } 
          else{
            ValidationUtils.verify(false,true,"Job Number is MisMatched");
          } 
                    
                  
        var monday = timesheetline.getItem(i).getText_2(3).OleValue.toString().trim();
          if(monday!=""){            
            ValidationUtils.verify(true,true,"Monday hour is Matched");
          } 
          else{
            ValidationUtils.verify(false,true,"Need to add Minimum hour in Monday");
          }                   
        var tuesday = timesheetline.getItem(i).getText_2(4).OleValue.toString().trim();
          if(tuesday!=""){
            ValidationUtils.verify(true,true,"Tuesday hour is Matched");
          } 
          else{
            ValidationUtils.verify(false,true,"Need to add Minimum hour in Tuesday");
          }           
        var Wednesday = timesheetline.getItem(i).getText_2(5).OleValue.toString().trim();
          if(Wednesday!=""){
            ValidationUtils.verify(true,true,"Wednesday hour is Matched");
          } 
          else{
            ValidationUtils.verify(false,true,"Need to add Minimum hour in Wednesday");
          }               
          var Thursday = timesheetline.getItem(i).getText_2(6).OleValue.toString().trim();
          if(Thursday!=""){
            ValidationUtils.verify(true,true,"Thursday hour is Matched");
          } 
          else{
            ValidationUtils.verify(false,true,"Need to add Minimum hour in Thursday");
          }            
           var Friday = timesheetline.getItem(i).getText_2(7).OleValue.toString().trim();
          if(Friday!=""){
            ValidationUtils.verify(true,true,"Friday hour is Matched");
          } 
          else{
            ValidationUtils.verify(false,true,"Need to add Minimum hour in Friday");
          } 
          

          
     }
      
      
} 