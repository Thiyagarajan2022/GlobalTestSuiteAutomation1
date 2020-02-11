//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "QueryExistingJob";
var Language = "";
Indicator.Show();
  
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile,Job_No,ProjectManager,RevenueMethod;



//getting data from datasheet
function getDetails(){
//Log.Message("excelName :"+workBook);
//Log.Message("sheet :"+sheetName);
ExcelUtils.setExcelName(workBook, sheetName, true);
//Log.Message(EnvParams.Opco)
comapany = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((comapany==null)||(comapany=="")){ 
ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");
}
Log.Message(comapany)
Job_name= ExcelUtils.getRowDatas("Job_name",EnvParams.Opco)
if((Job_name==null)||(Job_name=="")){ 
ValidationUtils.verify(false,true,"Job Name is Needed to Create a Job");
}
Log.Message(Job_name)

Job_No= ExcelUtils.getRowDatas("Job_No",EnvParams.Opco)
if((Job_No==null)||(Job_No=="")){ 
ValidationUtils.verify(false,true,"Job_No is Needed to Create a Job");
}
Log.Message(Job_No)
Job_group = ExcelUtils.getRowDatas("Job_group",EnvParams.Opco)
if((Job_group==null)||(Job_group=="")){ 
ValidationUtils.verify(false,true,"Job Group is Needed to Create a Job");
}
Log.Message(Job_group)
Job_Type = ExcelUtils.getRowDatas("Job_Type",EnvParams.Opco)
if((Job_Type==null)||(Job_Type=="")){ 
ValidationUtils.verify(false,true,"Job Type Number is Needed to Create a Job");
}
Log.Message(Job_Type)
department = ExcelUtils.getRowDatas("Department",EnvParams.Opco)
if((department==null)||(department=="")){ 
ValidationUtils.verify(false,true,"Department Number is Needed to Create a Job");
}
buss_unit = ExcelUtils.getRowDatas("BusinessUnit",EnvParams.Opco)
if((buss_unit==null)||(buss_unit=="")){ 
ValidationUtils.verify(false,true,"BusinessUnit Number is Needed to Create a Job");
}

ProjectManager =ExcelUtils.getRowDatas("Project_Manager",EnvParams.Opco)
if((ProjectManager==null)||(ProjectManager=="")){ 
ValidationUtils.verify(false,true,"ProjectManager is Needed to Create a Job");

}

TemplateNo = ExcelUtils.getRowDatas("Template",EnvParams.Opco)
if((TemplateNo==null)||(TemplateNo=="")){ 
ValidationUtils.verify(false,true,"Template Number is Needed to Create a Job");
}

RevenueMethod = ExcelUtils.getRowDatas("Revenue_Method",EnvParams.Opco)
//Log.Message(RevenueMethod);
if((RevenueMethod==null)||(RevenueMethod=="")){ 
ValidationUtils.verify(false,true,"RevenueMethod is Needed to Create a Job");
}



Product = ExcelUtils.getRowDatas("Product",EnvParams.Opco)
  if((Product=="")||(Product==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  Product = ReadExcelSheet("Product Number",EnvParams.Opco,"Data Management");
  }
if((Product==null)||(Product=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Create a Job");
}

ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//OpCoFile=ExcelUtils.getRowData1("OpCo File")
//if((OpCoFile==null)||(OpCoFile=="")){ 
//ValidationUtils.verify(false,true,"OpCoFile is Needed to Create a Job");
//}
}



//Validating Created Job is available
function GoToJob() {
  
      ReportUtils.logStep_Screenshot("Start Validating Job Details");
 TextUtils.writeLog("Navigate to Job"); 
  var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  aqUtils.Delay(2000, Indicator.Text);
//  Delay(2000);

  var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  companyFilter.forceFocus();
  companyFilter.setVisible(true);
  companyFilter.ClickM();
//  table.Child(0).setText("^a[BS]");
  table.Child(0).setText(comapany);
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
//  Delay(1000);
var jobNoField = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.firstcell
jobNoField.Click();
  jobNoField.setText(Job_No);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();
  
    
      ReportUtils.logStep_Screenshot("Enter Job Details");
 TextUtils.writeLog("Navigate to Job"); 

  job.setText(Job_name);
  aqUtils.Delay(7000, Indicator.Text);
//  Delay(7000);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(Job_name)){ 
//  var ExlArray = getExcelData("Validate_Company",EnvParams.Opco)  
//   var name =  LogReport_name(ExlArray,comapany,Job_group);
//      var notepadPath = Project.Path+"RegressionLogs\\"+EnvParams.instanceData+"\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+name+".txt";
////      Log.Message("Notepad :"+notepadPath)
//      TextUtils.writeDetails(notepadPath,"Job Number ",table.getItem(v).getText_2(2).OleValue.toString().trim());
     // ExcelUtils.setExcelName(workBook,"Data Management", true);
    //  ExcelUtils.WriteExcelSheet("Job Number",EnvParams.Opco,"Data Management",table.getItem(v).getText_2(2).OleValue.toString().trim())
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
Delay(8000);
  ValidationUtils.verify(flag,true,"Job Created is available in system");
  ValidationUtils.verify(true,true,"Job Number :"+table.getItem(v).getText_2(2).OleValue.toString().trim());
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  ReportUtils.logStep_Screenshot("");
  closeFilter.Click();
  aqUtils.Delay(8000, Indicator.Text);
//  Delay(8000);
//    var Home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//    Home.Click();
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  }
  
  var JobInfoTab = Aliases.ObjectGroup.JobInfoTab;
  JobInfoTab.Click();
        ReportUtils.logStep_Screenshot("Verifying Job Details");
 TextUtils.writeLog("Verifying Job Details"); 
   Delay(8000);
   
   
   var product = Aliases.ObjectGroup.JobGroupLabel;
   Log.Message(product.getText().trim())
   Sys.HighlightObject(product);
   var company =Aliases.ObjectGroup.CompanyName;
   Log.Message(company.getText().trim())
         Sys.HighlightObject(company);
   var revenueMethod = Aliases.ObjectGroup.RevenueMethod;
   Log.Message(revenueMethod.getText().trim())
      Sys.HighlightObject(revenueMethod);
   var jobgroup = Aliases.ObjectGroup.JobGroup;
     Log.Message(jobgroup.getText().trim())
        Sys.HighlightObject(jobgroup);
     var jobType =Aliases.ObjectGroup.JobTypeCode;
          Log.Message(jobType.getText().trim())
        Sys.HighlightObject(jobType);
     var jobDepartment = Aliases.ObjectGroup.JobDepartment;
          Log.Message(jobDepartment.getText().trim())
        Sys.HighlightObject(jobDepartment);
     var jobBusinessUnit =Aliases.ObjectGroup.BusinessUnit;
          Log.Message(jobBusinessUnit.getText().trim())
             Sys.HighlightObject(jobBusinessUnit);
     var projectManager =Aliases.ObjectGroup.ProjectManager;
           Log.Message(projectManager.getText().trim())
         Sys.HighlightObject(projectManager);
      
//        if(product.getText().trim().contains(Product))
//        {
//          ReportUtils.logStep("INFO", "Product Matched Correctly");
//          ReportUtils.logStep_Screenshot("");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//          ValidationUtils.verify(false,true,"Product No Does not Match");
//        }
        
        
               if(product.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "Product is present");
            ValidationUtils.verify(true,true,"Product is present");
              Log.Checkpoint("Product is present")
        }
        else{
           ValidationUtils.verify(false,true,"product is present");
        }
        
        
        
     //
        
//          if(company.getText().trim().contains(comapany))
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"Company No Does not Match");
//        }
        
        
               if(company.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "Company is present");
            ValidationUtils.verify(true,true,"Company is present");
              Log.Checkpoint("Company is present")
        }
        else{
           ValidationUtils.verify(false,true,"Company info is not present");
        }
        
      //  revenueMethod.gett
        
//           if(revenueMethod.getText().trim()==RevenueMethod)
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"revenueMethod Does not Match");
//           }
           
           
                  if(revenueMethod.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "RevenueMethod is present");
            ValidationUtils.verify(true,true,"RevenueMethod is present");
              Log.Checkpoint("RevenueMethod is present")
        }
        else{
           ValidationUtils.verify(false,true,"RevenueMethod is present");
        }         
   
        
//            if(jobgroup.getText().trim()==Job_group)
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"jobgroup Does not Match");
//        }
        
        
               if(jobgroup.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "jobgroup is present");
            ValidationUtils.verify(true,true,"jobgroup is present");
              Log.Checkpoint("jobgroup is present")
        }
        else{
           ValidationUtils.verify(false,true,"jobgroup is present");
        }
           
        
//            if(jobDepartment.getText().trim()==department)
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"jobDepartment Does not Match");
//        }
        
               if(jobDepartment.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "jobDepartment is present");
            ValidationUtils.verify(true,true,"jobDepartment is present");
              Log.Checkpoint("jobDepartment is present")
        }
        else{
           ValidationUtils.verify(false,true,"jobDepartment is present");
        }
           
        
//            if(jobType.getText().trim()==Job_Type)
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"jobType Does not Match");
//        }
        
               if(jobType.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "jobType is present");
            ValidationUtils.verify(true,true,"jobType is present");
              Log.Checkpoint("jobType is present")
        }
        else{
           ValidationUtils.verify(false,true,"jobType is present");
        }
           
        
//            if(jobBusinessUnit.getText().trim()==buss_unit)
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"jobBusinessUnit Does not Match");
//        }
        
               if(jobBusinessUnit.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "jobBusinessUnit is present");
            ValidationUtils.verify(true,true,"jobBusinessUnit is present");
              Log.Checkpoint("jobBusinessUnit is present")
        }
        else{
           ValidationUtils.verify(false,true,"jobBusinessUnit is present");
        }
        
             
//            if(projectManager.getText()==ProjectManager)
//        {
//           ReportUtils.logStep("INFO", "Product Matched Correctly");
//            ValidationUtils.verify(true,true,"Product Matched Correctly");
//        }
//        else{
//           ValidationUtils.verify(false,true,"projectManager Does not Match");
//        }
        
        
               if(projectManager.getText().trim()!="")
        {
           ReportUtils.logStep("INFO", "projectManager is present");
            ValidationUtils.verify(true,true,"projectManager is present");
              Log.Checkpoint("projectManager is present")
        }
        else{
           ValidationUtils.verify(false,true,"projectManager is present");
        }
           
  }      
}

//Go To Job from Menu
function goToJobMenuItem(){

aqUtils.Delay(5000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
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

var mainlist = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var main;
for(var id=0;id<mainlist;id++){
main = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
if(main.Child(id).isVisible())
if(main.Child(id).ChildCount==1)
if(main.Child(id).Child(0).Name.indexOf("Composite")!=-1){

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot("");
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}
}
}

}
//Delay(5000); 
ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
 
}

//Main Function
function QueryExistingJob() {
//Indicator.PushText("waiing for window to open");
//var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//
//menuBar.Click();
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
//if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
//    Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x58); //X 
//    Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//    Sys.Desktop.KeyUp(0x58);
//Restart.login(Project_manager);
//  
//}
//aqTestCase.Begin("Job Creation", "zfj://CH1-67");
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "QueryExistingJob";
Language = "";
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
comapany,Job_group,Job_Type,department,buss_unit,TemplateNo,Product,Job_name,Project_manager,OpCoFile,Job_No,ProjectManager,RevenueMethod ="";


Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Log.Message(EnvParams.Opco)
Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Querying Existing Job started::"+STIME);
getDetails();
goToJobMenuItem();      
GoToJob();
WorkspaceUtils.closeAllWorkspaces();
//aqTestCase.End();

}


function getExcel(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf(",")!=-1){
     var excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
}



function getExcelData(rowidentifier,column) { 
var temp = ""
//var array = "Validate_EmployeeCategories";
//var Opco = "1307"
var excelData = [];
//Log.Message("Execution completed,sending result to excel book , FileName:"+excelName+"sheetname:"+sheet);
  var app = Sys.OleObject("Excel.Application");
//  app.Visible = "True";
  var curArrayVals = [];  
//  Log.Message(workBook)
//  Log.Message(sheetName)
//  Log.Message(rowidentifier)
//  Log.Message(column)
  var book = app.Workbooks.Open(workBook);
  var sheet = book.Sheets.Item(sheetName);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;
//  Log.Message(columnCount);
//  Log.Message(rowCount);
  var arrays={};
  var idx =0;
  var col =0;
  var row = 0;
  for(var k = 1; k<=columnCount;k++){
  if(sheet.Cells.Item(1, k).Text.toString().trim().toUpperCase()==column.toUpperCase()){
  col = k;
  }
  }
  var rowStatus = false;
  for(var k = 1; k<=rowCount;k++){
//  Log.Message(sheet.Cells.Item(k, 1).Text.toString().trim())
  if(sheet.Cells.Item(k, 1).Text.toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
  row = k;
  rowStatus = true;
  }
  }
//  Log.Message(col)
//  Log.Message(row);
  if(rowStatus){ 
   temp = sheet.Cells.Item(row,  col).Text;
//   Log.Message(temp)
  }
  
  
// book.Save();
 app.Quit();
 
 
 if(temp.indexOf(",")!=-1){ 
//       Log.Message(temp)
      excelData =  temp.split(",");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     

 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);

 return excelData;
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf("*")!=-1){
     var excelData =  temp.split("*");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
     DDT.CloseDriver(xlDriver.Name);
 for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
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
Log.Message(compStatus);
return compStatus
}
