//USEUNIT ActionUtils
//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ObjectUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


function MainFunction(){ 

// Starting HTML File
var report = JavaClasses.com_relevantcodes_extentreports.ExtentReports.newInstance(Project.Path +"KT_Sample\\KT_Session_1.html");
var test = report.startTest("Sample Test","TestCase Description");
test.log(logStatus.INFO,"TestCase is Started");
test.log(logStatus.PASS,"Step 2");
test.log(logStatus.WARNING,"Step 3");
test.log(logStatus.FAIL,"Step 3");


// End of HTML File
report.endTest(test);

var test = report.startTest("Sample Test 2"," 2 TestCase Description");
test.log(logStatus.INFO,"2 TestCase is Started");
test.log(logStatus.PASS,"Step 2");
test.log(logStatus.WARNING,"Step 3");



// End of HTML File
report.endTest(test);

report.flush();


}



function SampleJob() {
// Starting HTML File
var report = JavaClasses.com_relevantcodes_extentreports.ExtentReports.newInstance(Project.Path +"KT_Sample\\KT_Session_1.html");
var test = report.startTest("Job Screen","Checking Job Screen fields");  

var WorkspceCLient = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
Sys.HighlightObject(WorkspceCLient);
WorkspceCLient.HoverMouse();
WorkspceCLient.DblClick();
test.log(logStatus.INFO,"WorkSpace fields is clicked");

ImageRepository.Spanish.Job_1.Click();

test.log(logStatus.INFO,"Jobs Image is clicked in Workspace Client");

var Jobs = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 4).SWTObject("Tree", "");
Sys.HighlightObject(Jobs);
Jobs.DblClickItem("|Jobs");

if(ImageRepository.Spanish.Loading.Exists()){ }

test.log(logStatus.PASS,"Jobs workspace is loaded completly");

var AllJobs = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
AllJobs.Click();
test.log(logStatus.PASS,"All Jobs is Clicked");
if(ImageRepository.Spanish.Loading.Exists()){ }

var newJob = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
newJob.Click();
test.log(logStatus.INFO,"New Jobs is Clicked");
if(ImageRepository.Spanish.Loading.Exists()){ }

var company = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
company.Click();


//0x11 // Ctrl Key
//0x47 // 'G' key

Sys.Desktop.KeyDown(0x11);
Sys.Desktop.KeyDown(0x47);
Sys.Desktop.KeyUp(0x47);
Sys.Desktop.KeyUp(0x11);

if(ImageRepository.Spanish.Loading.Exists()){ }

var ComID = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
ComID.Keys("1221");

var search = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
search.Click();
Delay(5000);
var table = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2)

for(var i=0;table.length;i++){ 
 if(table.getItem(i).getText(0).OleValue.toString().trim()=="1221")  {
   table.Keys("[Down]");
   test.log(logStatus.PASS,"Company Number is availble and clicked in maconomy")
   break;
 }else{ 
   table.Keys("[Down]")
 }

}

var Okay = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
Okay.Click();

var DropDown = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
DropDown.Click();

aqUtils.Delay(5000,"Waiting for DropDown")
Delay(5000);

ScrollBar = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);

for(var i=0;ScrollBar.length;i++){ 
 if(ScrollBar.getItem(i).getText(0).OleValue.toString().trim()=="Client Billable")  {
   ScrollBar.Keys("[Enter]");
   test.log(logStatus.PASS,"Job Group is Selected in maconomy")
   break;
 }else{ 
   ScrollBar.Keys("[Down]")
 }

}

var Cacel = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
Cacel.Click();

test.log(logStatus.INFO,"Job Screen Checking is Completed");


report.endTest(test);
report.flush();




}




