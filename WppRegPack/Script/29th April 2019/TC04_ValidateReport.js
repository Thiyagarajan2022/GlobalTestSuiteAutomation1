//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT PdfUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "MPLValidation";
var temporary="";
var JobName_strIndex;  
var ExcelDetails = [];
var JobDetails = [];
var JobNo = "";
var Job_group ="";
var Job_Type ="";
var department ="";
var buss_unit = "";
var TemplateNo ="";
var Product ="";
var Product_name ="";
var Client ="";
var Client_Name ="";
var Brand ="";
var Brand_Name ="";
var Job_name="";
var Project_manager ="";
var temp="";

function SOXexcel(sheetName,start){ 
var Arrayss = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, false);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
     Arrayss[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrayss;
}



/*


function excel(){ 

var Arrayss = [];
var xlDriver = DDT.ExcelDriver(Project.Path+"\\DataTables"+"\\Job_Detail - 1707200094.xlsx", "Job Detail", true);
var id =0;
var colsList = [];
//Log.Message(DDT.CurrentDriver.ColumnCount);
   xlDriver.Next();
   xlDriver.Next();
   for(var idx=0;idx<xlDriver.ColumnCount;idx++){   
     colsList[idx] = xlDriver.ColumnName(idx);
//  Log.Message(colsList[idx]);   
   }
//   xlDriver.Next();

     while (!DDT.CurrentDriver.EOF()) {
     Log.Message( colsList.length );
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
//      if(idx!=colsList.length-1)
//      temp = temp+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
     if(temp.length!=6){
     Arrayss[id]=temp;
//     Log.Message(temp)
     }
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return Arrayss;
}



function getExcelData(){
    excelData =[];  
    var colsList = [];
    var xlDriver = DDT.ExcelDriver(Project.Path+"\\DataTables"+"\\Job_Detail - 1707200094.xlsx","Job Detail", true)
    for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
    var data = "";
    var i = 0;
    while (!DDT.CurrentDriver.EOF()) {
    data = "";
    for(var idx=0;idx<colsList.length;idx++){ 
    try{
       data = data + xlDriver.Value(colsList[idx]).toString().trim() + "*";
       excelData[i] = data;
       
       }
       catch(err)
       {
       data = data +"*";
       excelData[i] = data;
       }
       
       }
      // Log.Message("EXCELDATA :"+excelData[i]);      
       i++;
     xlDriver.Next();
    }
    
    DDT.CloseDriver(xlDriver.Name);
   // Log.Message("completed reading excel data, data length::"+excelData.length);
    return excelData;
  
}


*/

var RecNo;
var JobDetails = "";
var CostFees = [];
var Estimate = [];
var Billing = [];
var Summary = [];
var jobStatus = 0;
// Posts data to the log (helper routine)
function ProcessData()
 {
  var Fldr, i;
  

var temp = "";  
  for(i = 0; i < DDT.CurrentDriver.ColumnCount; i++){ 
temp = temp+aqConvert.VarToStr(DDT.CurrentDriver.Value(i));
}

if(temp.indexOf("Costs & Fees")==0){ 
  jobStatus = 1;
  RecNo = 0;
}
if(temp.indexOf("Estimate History")==0){ 
  jobStatus = 2;
  RecNo = 0;
}
if(temp.indexOf("Billing")==0){ 
  jobStatus = 3;
  RecNo = 0;
}
if(temp.indexOf("Summary")==0){ 
  jobStatus = 4;
  RecNo = 0;
}


if(jobStatus==0){

  var temporary1="";

  var JobName_EndIndex;
if(RecNo==2){ 
  temporary = temp;
  JobName_strIndex = temp.indexOf("Client Name");

}
else if(RecNo==3){ 
  temporary1 = temp;
  JobName_EndIndex = temp.indexOf("Brand Code");
  if(JobName_EndIndex==0){ 
    JobDetails = JobDetails+temporary+temporary1;
  }
  if(JobName_EndIndex>0){ 
//    Log.Message(temporary)
//    Log.Message(JobName_strIndex);
//    Log.Message(temporary.substring(0,JobName_strIndex));
//    Log.Message(temporary1.substring(0,JobName_EndIndex));
//    Log.Message(temporary.substring(JobName_strIndex));
//    Log.Message(temporary1.substring(JobName_EndIndex));
    JobDetails = JobDetails+temporary.substring(0,JobName_strIndex)+temporary1.substring(0,JobName_EndIndex)+temporary.substring(JobName_strIndex)+temporary1.substring(JobName_EndIndex);
  }
}
else{ 
 JobDetails = JobDetails+temp; 
}
}
if(jobStatus==1){
CostFees[RecNo] = temp;
}
if(jobStatus==2){
Estimate[RecNo] = temp;
}
if(jobStatus==3){
Billing[RecNo] = temp;
}
if(jobStatus==4){
Summary[RecNo] = temp;
}










// Log.Message(temp); 
//  Log.PopLogFolder(); 
  RecNo = RecNo + 1; 
 }
  
// Creates the driver (main routine)
function TestDriver()
 {
  var Driver;
  
  // Creates the driver
  // If you connect to an Excel 2007 sheet, use the following method call:
  Driver = DDT.ExcelDriver(Project.Path+"\\"+JobDetails[2]+"\\"+JobDetails[3], "Job Detail"); 
  
  // Iterates through records
  RecNo = 0;
  while (! Driver.EOF() ) 
  {
    ProcessData(); // Processes data
    Driver.Next(); // Goes to the next record
  }
  
  // Closing the driver
  DDT.CloseDriver(Driver.Name);
//  ExcelDetails = SOXexcel(sheetName,1);

    Log.Message(JobDetails);

  ValidationUtils.verify(JobDetails.indexOf("Job No."+ExcelDetails[0])!=-1,true,"Job Number is available");
  
  ValidationUtils.verify(JobDetails.indexOf("Job"+ExcelDetails[1])!=-1,true,"Job Name is available");
  
  ValidationUtils.verify(JobDetails.indexOf("Job Type"+ExcelDetails[2])!=-1,true,"Job_Type is available");

//  ValidationUtils.verify(JobDetails.indexOf(ExcelDetails[3])!=-1,true,"Department is available");

//  ValidationUtils.verify(JobDetails.indexOf(ExcelDetails[4])!=-1,true,"Business Unit is available");

  ValidationUtils.verify(JobDetails.indexOf("Product Code"+ExcelDetails[3])!=-1,true,"Product Number is available");

  ValidationUtils.verify(JobDetails.indexOf("Product Name"+ExcelDetails[4])!=-1,true,"Product Name is available");

  ValidationUtils.verify(JobDetails.indexOf("Client Code"+ExcelDetails[5])!=-1,true,"Client Number is available");

  ValidationUtils.verify(JobDetails.indexOf("Client Name"+ExcelDetails[6])!=-1,true,"Client Name is available");

  ValidationUtils.verify(JobDetails.indexOf("Brand Code"+ExcelDetails[7])!=-1,true,"Brand Number is available");

  ValidationUtils.verify(JobDetails.indexOf("Brand Name"+ExcelDetails[8])!=-1,true,"Brand Name is available");

  ValidationUtils.verify(JobDetails.indexOf("Project Manager"+ExcelDetails[9])!=-1,true,"Project Manager Number is available");


  
  
//  Log.Message("==========================================================================");
//  for(var i=0;i<CostFees.length;i++){ 
//    Log.Message(CostFees[i]);
//  }
//  Log.Message("==========================================================================");
//  for(var i=0;i<Estimate.length;i++){ 
//    Log.Message(Estimate[i]);
//  }
//  Log.Message("==========================================================================");
//  for(var i=0;i<Billing.length;i++){ 
//    Log.Message(Billing[i]);
//  }
//  Log.Message("==========================================================================");
//  for(var i=0;i<Summary.length;i++){ 
//    Log.Message(Summary[i]);
//  }
//  Log.Message("==========================================================================");
  
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
var jobSubItem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 4).SWTObject("Tree", "");
//  var jobSubItem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "")
jobSubItem.DblClickItem("|Jobs"); 
Delay(5000); 
ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
 
}

//Go To Job To Validate Report
function GoToJob() {
var newJobBtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 1);

var Add_Visible = true;
while(Add_Visible){
if(newJobBtn.isEnabled()){
Delay(2000);
Add_Visible = false;  
}
}
  var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  Delay(2000);

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
//  table.Child(0).setText(comapany);
  Delay(1000);
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Delay(1000);
  Sys.Desktop.KeyDown(0x09);
  Delay(1000);
  var jobNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  jobNo.Click();

  jobNo.setText(JobDetails[0]);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();

  job.setText(JobDetails[1]);
  Delay(3000);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if((table.getItem(v).getText_2(3).OleValue.toString().trim()==(JobDetails[1]))&&(table.getItem(v).getText_2(2).OleValue.toString().trim()==(JobDetails[0]))){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  closeFilter.Click();
  Delay(8000);
  var count = true;
//    var Home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//    Home.Click();
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  info.Click();
  count=false;
  }
  Delay(3000);
  
  
  var jobNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3);
  jobNumber.Click();
  ExcelDetails[0] = jobNumber.getText();
  var JobName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
  JobName.Click();
  ExcelDetails[1] = JobName.getText();  
//  Job_group = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
//  Job_group.Click();
//  ExcelDetails[2] = Job_group.getText();
  Job_Type = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
  Job_Type.Click();
  ExcelDetails[2] = Job_Type.getText();
  Product = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  Product.Click();
  temp = Product.getText().OleValue.toString().trim();
  ExcelDetails[3] = temp.substring(temp.indexOf(" (")+2,temp.indexOf(")"));
  ExcelDetails[4] = temp.substring(0,temp.indexOf(" ("));
  
  Client = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
  Client.Click();
  temp = Client.getText().OleValue.toString().trim();
  ExcelDetails[5] = temp.substring(temp.indexOf(" (")+2,temp.indexOf(")"));
  ExcelDetails[6] = temp.substring(0,temp.indexOf(" ("));
  
  Brand = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
  Brand.Click();
  temp = Brand.getText().OleValue.toString().trim();
  ExcelDetails[7] = temp.substring(temp.indexOf(" (")+2,temp.indexOf(")"));
  ExcelDetails[8] = temp.substring(0,temp.indexOf(" ("));
  
  Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
  Project_manager.Click();
  ExcelDetails[9] = Project_manager.getText();
  
  
  
 /* 
  
  var Templete_Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Blanket_invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  var estimation = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Amount_Registrations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Invocing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", ""); 
  var TimeReg = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  checkmark = false;
  if(Templete_Job.getSelection()){ 
  Templete_Job.Click();
    Log.Message("Templete_Job is UnChecked")
    checkmark = true;
  }
  if(Blanket_invoice.getSelection()){ 
  Blanket_invoice.Click();
    Log.Message("Blanket_invoice is UnChecked")
    checkmark = true;
  }
  if(Amount_Registrations.getSelection()){ 
  Amount_Registrations.Click();
    Log.Message("Amount_Registrations is UnChecked")
    checkmark = true;
  }
  if(Invocing.getSelection()){ 
  Invocing.Click();
    Log.Message("Invocing is UnChecked")
    checkmark = true;
  }
  if(TimeReg.getSelection()){ 
  TimeReg.Click();
    Log.Message("TimeRegistration is UnChecked")
    checkmark = true;
  }
  if(estimation.getSelection()){ 
  estimation.Click();
    Log.Message("Estimating is UnChecked")
    checkmark = true;
  }

  if(checkmark){ 
      
    Delay(3000);
    var save_change = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
    save_change.Click();
    Log.Message("Changes is Saved");
    Delay(5000);
  }
  ReportUtils.logStep("INFO", "Job is Saved Pending for Approval");
  */
  var filter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
  filter.Click();
}
}

//Main Function
function ValidateJobReport(){ 
  gotoReporting();
  Delay(9000);
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  WorkspaceUtils.closeAllWorkspaces();
  goToJobMenuItem();
  JobDetails = SOXexcel(sheetName,1);
  ExcelDetails = [];
  GoToJob();
  WorkspaceUtils.closeAllWorkspaces();
  TestDriver();
}


//Go to Report from Menu 
function goToReportItem(){

//   var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "");
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.DblClick();
 if(ImageRepository.ImageSet.Reporting.Exists()){
 ImageRepository.ImageSet.Reporting.Click();// GL
}
else if(ImageRepository.ImageSet.Reporting1.Exists()){
ImageRepository.ImageSet.Reporting1.Click();
}
else{
ImageRepository.ImageSet.Reporting2.Click();
}
var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|Reports");
} 
}
Delay(5000); 
//ReportUtils.logStep("INFO", "Moved to Reports from Reports Menu");
 
}


//Exporting Report from Maconomy
function goToReport(){ 
//Delay(5000);
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
  job.Click();
  Delay(3000);
  var jobDetail = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McLinkLabelWidget", "", 2).SWTObject("McTextWidget", "");
  Sys.HighlightObject(jobDetail);
  jobDetail.Click();
  Delay(4000);
  
  
}
function gotoReporting(){ 
  goToReportItem();
  goToReport();
}