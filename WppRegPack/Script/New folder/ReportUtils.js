var report,test,reportConsolidated,testConsolidated;
var file_path;

function createReport(filePath,fileName)
{
file_path = filePath;
report = JavaClasses.com_relevantcodes_extentreports.Extent_Craft.newInstance(); 
report.initialize(filePath,fileName);
}

function createConsolidatedReport(filePath,fileName)
{
file_path = filePath;
reportConsolidated = JavaClasses.com_relevantcodes_extentreports.Extent_Craft.newInstance(); 
reportConsolidated.initialize(filePath,fileName);

}

function createTest(testName,testDesc)
{
test = report.createTest(testName,testDesc);
testConsolidated = reportConsolidated.createTest(testName,testDesc);

}

function logStep(result,stepName, stepDesc="")
{
//var image = "image_"+aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y_%H.%M.%S")+".png"
//if(result=="FAIL"){
//aqFileSystem.CreateFolder(file_path+"images");
//  Sys.Desktop.Picture().SaveToFile(file_path+"images\\"+image);
//  
//}
//report.logStep(result,stepName, stepDesc, file_path+"images\\"+image );

//Log.Message(stepName);
//Log.Message(result);
report.logStep(result,stepName, "info");
reportConsolidated.logStep(result,stepName, stepDesc);
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
