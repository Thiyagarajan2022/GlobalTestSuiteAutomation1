//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "Sub Jobs";

//var workBook = "C:\\Users\\741496\\Desktop\\SIT Test Samples.xlsx";
//var sheetName ="Sub Jobs";
var Arrays = [];
var count = true;
var checkmark = false;
var company ="";
var Job_group ="";
var Job_Type ="";
var department ="";
var buss_unit = "";
var TemplateNo ="";
var Product ="";
var Job_name="";
var Sub_Job="";
var Project_manager ="";

  function GoToSubJob() {
    var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

    Delay(3000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    Delay(1000);

    var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
    SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
    SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
    SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
    SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
    SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    companyFilter.forceFocus();
    companyFilter.setVisible(true);
    companyFilter.Click();
    table.Child(0).setText("^a[BS]");
    table.Child(0).setText(company);
    Delay(1000);
    Sys.Desktop.KeyDown(0x09); // Press Ctrl
    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    Delay(1000);
    Sys.Desktop.KeyDown(0x09);
    
    Sys.Desktop.KeyUp(0x09);
    var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3); 
    job.Click();
    Delay(4000)
    table.Child(2).forceFocus();
    table.Child(2).setVisible(true);
//    table.Child(2).setText(dataJobName);
    table.Child(2).setText(Job_name);
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(8000);
    
    if(count){
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
    ref.Refresh();
    var subjob = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 10)
    subjob.Click()
    Delay(4000);
    count=false;
    }
    
      var createSubJobbtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 6);
      Sys.HighlightObject(createSubJobbtn);
      createSubJobbtn.Click();
 
     
    Delay(2000)
    var screen = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job");
//    Log.Message("Create New SubJob window is Opened")
    
    var jobgroup = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
    jobgroup.Click()
    Delay(2000)
    jobgroup.Keys(Job_group);
    Delay(5000)

    
    var jobtype =Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
    jobtype.Click();    
    jobtype.setText(Job_Type);
    Delay(5000)
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(2000)
    
    var depart = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
    depart.Click();
    depart.setText(department);
    Delay(5000)
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(2000)
    
    var business = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(buss_unit!=""){
  business.Click();
  WorkspaceUtils.SearchByValue(business,"Business Unit",buss_unit);
//
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x11);
//  Sys.Desktop.KeyDown(0x47);
//  Sys.Desktop.KeyUp(0x11);
//  Sys.Desktop.KeyUp(0x47);
//  Delay(3000);
//  var code = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//  code.setText(buss_unit);
//  Delay(3000);
//  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
//  Sys.HighlightObject(serch);
//  serch.Click();
//  Delay(5000);
//  var table = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
//  Sys.HighlightObject(table);
//  var itemCount = table.getItemCount();
//  if(itemCount>0){ 
//  for(var i=0;i<itemCount;i++){
//    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==buss_unit){ 
//     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
//        OK.Click();
//          
//    }
//    else{ 
//      Sys.Desktop.KeyDown(0x28);
//      Sys.Desktop.KeyUp(0x28);
//      if(i==itemCount-1){ 
//        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//        cancel.Click();
//        Delay(1000);
//        BussUnit.setText("");
//      }
//    }
//      
//    }
//  }
//  else { 
//    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//    cancel.Click();
//    Delay(1000);
//    BussUnit.setText("");
//  } 
            }
else{ 
    ValidationUtils.verify(false,true,"Business Unit Number is Needed to Create a Job");
    }  

    var template = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
    template.Click()    
    template.setText(TemplateNo);
    Delay(5000)
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(2000)
    
    var jobname = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
    jobname.click()
    jobname.setText(Sub_Job);
    Delay(3000)
 
    
    var project = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
    project.Click()
    project.Keys("^a[BS]");
    project.setText(Project_manager);
    Delay(3000)    
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(3000);
    
    
    var subjobbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
    Sys.HighlightObject(subjobbtn);
    if(subjobbtn.isEnabled()){
      subjobbtn.Click();
      Delay(3000)
      ValidationUtils.verify(true,true,"SubJob is Created");
      } 
    else{
      Log.Warning("Create Button is not Visible");
     ValidationUtils.verify(false,true,"SubJob is Not Created");
      var cancelbtn = Sys.Process("Maconomy").SWTObject("Shell", "Create Sub Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
    Sys.HighlightObject(cancelbtn);
    Delay(2000)
    cancelbtn.Click()
      Delay(3000)      
    } 
    
        
    var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2); 
    closefilter.Click();
    
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
    table.Click();      
    table.setText(Sub_Job);
    Delay(5000)
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(3000)
    
//    var OK = Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Job").SWTObject("Composite", "", 2).SWTObject("Button", "Yes");
//    var cancel =  Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Job").SWTObject("Composite", "", 2).SWTObject("Button", "No");
//    cancel.Click();
//    Delay(5000);  
    
    }

    
    function gotoinformation(){
   
    var information = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    information.Click();
    count=false;
    Delay(6000)
    
    var Amount_Registrations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    var Invocing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", ""); 
    var TimeReg = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    var Trmplete_Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    var Blanket = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    var estimating = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    checkmark = false;
    if(Trmplete_Job.getSelection()){ 
    Trmplete_Job.Click();
      Log.Message("Trmplete_Job is UnChecked")
      checkmark = true;
    }
    if(Blanket.getSelection()){ 
    Blanket.Click();
      Log.Message("Blanket is UnChecked")
      checkmark = true;
    }
    if(estimating.getSelection()){
      estimating.Click();
      Log.Message("estimating is Unchecked")
      checkmark = true;
    } 
    
    if(TimeReg.getSelection()){ 
    TimeReg.Click();
      Log.Message("TimeReg is UnChecked")
      checkmark = true;
    }
    if(Amount_Registrations.getSelection()){ 
    Amount_Registrations.Click();
      Log.Message("Amount_Registrations is UnChecked")
      checkmark = true;
    }
    if(Invocing.getSelection()){
      Invocing.Click();
      Log.Message("Invocing is Unchecked")
      checkmark = true;
    } 
    
    Delay(3000)
    if(checkmark){
    var savebtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4); 
    Sys.HighlightObject(savebtn);
    savebtn.Click();
    }
    Delay(1000);
    ReportUtils.logStep("INFO", "SubJob has Created and Available in Table ");
    
//    var popup = Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Information");
//    
//    if(savebtn){
//    var yes = Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Information").SWTObject("Composite", "", 2).SWTObject("Button", "Yes");
//    Sys.HighlightObject(yes);
//    yes.Click();
//    }
//    else
//    {
//    var no = Sys.Process("Maconomy").SWTObject("Shell", "Jobs - Information").SWTObject("Composite", "", 2).SWTObject("Button", "No");
//    Sys.HighlightObject(no);
//    no.Click();
//    }
    
    }



function goToJobMenuItem(){
//   var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "");
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
      
      
var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|Jobs");
}
}
      Delay(6000);
//    var jobSubItem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 4).SWTObject("Tree", "");
//    jobSubItem.DblClickItem("|Jobs"); 
//    Delay(5000);
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
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
     Arrays[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}


function runAll(){ 
    excel();
    goToJobMenuItem();
      for(var i=0;i<Arrays.length;i++){
        var splitArray = Arrays[i].split("*");
          company =splitArray[0];
          Job_name=splitArray[1];
          Job_group =splitArray[2];
          Job_Type =splitArray[3];
          department =splitArray[4];
          buss_unit = splitArray[5];
          TemplateNo =splitArray[6];
          Sub_Job=splitArray[7];
          Project_manager =splitArray[8];
    GoToSubJob();
    gotoinformation();
     }
    closeAllWorkspaces();
}



