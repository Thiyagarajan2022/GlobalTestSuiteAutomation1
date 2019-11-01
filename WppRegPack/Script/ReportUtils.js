//USEUNIT TestRunner
var report,test,reportConsolidated,testConsolidated;
var file_path;
var testExe = "";
var ig = 0;
var logStatus = JavaClasses.com_relevantcodes_extentreports.LogStatus;
function createReport(filePath,fileName)
{
file_path = filePath;
//Log.Message(filePath+fileName)
report = JavaClasses.com_relevantcodes_extentreports.ExtentReports.newInstance(filePath+fileName);


}

function createConsolidatedReport(filePath,fileName)
{
file_path = filePath;
//Log.Message(filePath+fileName)
reportConsolidated = JavaClasses.com_relevantcodes_extentreports.ExtentReports.newInstance(filePath+fileName+".html");
}

function createTest(testName,testDesc)
{
testExe = testName;
ig = 0;
test = report.startTest(testName,testDesc);
testConsolidated = reportConsolidated.startTest(testName,testDesc);
}


function logStep(result,stepName, stepDesc="")
{
if(result.toUpperCase()=="INFO"){
//Log.Message(stepName)
test.log(logStatus.INFO,stepName);
testConsolidated.log(logStatus.INFO,stepName);
}
if(result.toUpperCase()=="PASS"){
test.log(logStatus.PASS,stepName);
testConsolidated.log(logStatus.PASS,stepName);
}
if(result.toUpperCase()=="FAIL"){

sFolder = file_path+"\\Screenshots\\"+TestRunner.testOpco+"_"+testExe.substring(0,testExe.indexOf(" "))+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
var img = sFolder+"Image_"+ig+".png";
Sys.Desktop.Picture().SaveToFile(img);
img = ".\\Screenshots\\"+TestRunner.testOpco+"_"+testExe.substring(0,testExe.indexOf(" "))+"\\"+"Image_"+ig+".png";
ig++;
test.log(logStatus.FAIL,stepName+test.addScreenCapture(img));
testConsolidated.log(logStatus.FAIL,stepName+testConsolidated.addScreenCapture(img));
ReportUtils.report.endTest(test);
ReportUtils.report.flush();

ReportUtils.reportConsolidated.endTest(testConsolidated);
ReportUtils.reportConsolidated.flush();
}
if(result.toUpperCase()=="WARNING"){
test.log(logStatus.WARNING,stepName);
testConsolidated.log(logStatus.WARNING,stepName);
}
  
}

function logStep_Screenshot(stepName)
{


sFolder = file_path+"\\Screenshots\\"+TestRunner.testOpco+"_"+testExe.substring(0,testExe.indexOf(" "))+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
var img = sFolder+"Image_"+ig+".png";
Sys.Desktop.Picture().SaveToFile(img);

img = ".\\Screenshots\\"+TestRunner.testOpco+"_"+testExe.substring(0,testExe.indexOf(" "))+"\\"+"Image_"+ig+".png";
ig++;
if((stepName!="")&&(stepName!=null)){
test.log(logStatus.INFO,stepName+test.addScreenCapture(img));
testConsolidated.log(logStatus.INFO,stepName+testConsolidated.addScreenCapture(img));
}
else
{ 
test.log(logStatus.INFO,test.addScreenCapture(img));
testConsolidated.log(logStatus.INFO,testConsolidated.addScreenCapture(img));
  
}
}

function screenshot()
{
 img = Sys.Desktop.FocusedWindow().Picture();
 return img;
}



function ReportLocator(TestCase)
{
var outFileName;
var filename = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d_%H%M");

Log.Event("Converting output to mht...");
outFileName = Project.ConfigPath + "Log\\" + filename + ".mht";
Log.Message(outFileName);
Log.SaveResultsAs(outFileName, 2); //converts output to mht

var pjtpath = Project.Path + "\ExecutionResults\\";

// Creates the folder and checks if it has been created successfully
if (aqFileSystem.CreateFolder(pjtpath) == 0)
// Creates a file in that folder
aqFile.Create(pjtpath);
else
//Log.Message("Could not create the folder " + pjtpath);

var reportPath = pjtpath + TestCase + aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y_%H.%M");//"%Y%m%d_%H%M");
aqFileSystem.CreateFolder(reportPath);

if(aqFile.Exists(reportPath + filename + ".mht")){
aqFile.Delete(reportPath + filename + ".mht");
}

//Copy the file elsewhere
var newPath = reportPath + "\\" +filename + ".mht";
//Log.Message(newPath);
if(aqFile.Move(outFileName, newPath)){
Log.Message("File moved");
} else {
Log.Warning("File was not moved!");
}
}

/*
if(aqFile.Exists("C:\\Users\\Administrator\\Desktop\\TC_Log\\" + filename + ".mht")){
aqFile.Delete("C:\\Users\\Administrator\\Desktop\\TC_Log\\" + filename + ".mht");
}

//Copy the file elsewhere
if(aqFile.Move(outFileName, "C:\\Users\\Administrator\\Desktop\\TC_Log\\" + filename + ".mht")){
Log.Message("File moved");
} else {
Log.Warning("File was not moved!");
}
}
*/

function d()
{
Log.Message(aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y_%H.%M.%S"));
}
