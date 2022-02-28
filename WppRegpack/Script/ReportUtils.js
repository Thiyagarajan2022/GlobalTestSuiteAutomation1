﻿//USEUNIT TestRunner
//USEUNIT TextUtils
var report,test,Dtest,Dreport,reportConsolidated,testConsolidated;
var file_path;
var file_name;
var testExe = "";
var ig = 0;
var logStatus = JavaClasses.com_relevantcodes_extentreports.LogStatus
var Dfile_path;
var Dfile_name;
var DStat = false;
function createReport(filePath,fileName)
{
file_path = filePath+fileName;
file_name = fileName;
//Log.Message(filePath+fileName)
report = JavaClasses.com_relevantcodes_extentreports.ExtentReports.newInstance(file_path+"\\"+fileName+".html");
}

function createDependencyReport(fileName)
{
Dfile_path = file_path+"\\"+fileName;
Dfile_name = fileName;
Log.Message(Dfile_path)
Dreport = JavaClasses.com_relevantcodes_extentreports.ExtentReports.newInstance(Dfile_path+"\\"+Dfile_name+".html");
}

function createConsolidatedReport(filePath,fileName)
{
file_path = filePath;
file_name = fileName;
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

function DependycreateTest(testName,testDesc)
{
testExe = testName;
ig = 0;
Dtest = Dreport.startTest(testName,testDesc);
test = report.startTest(testName,testDesc);
testConsolidated = reportConsolidated.startTest(testName,testDesc);
}


function logStep(result,stepName, stepDesc="")
{
if(result.toUpperCase()=="INFO"){
//Log.Message(stepName)
//TextUtils.writeLog(stepName);
if(DStat){ 
Dtest.log(logStatus.INFO,stepName);
}
test.log(logStatus.INFO,stepName);
testConsolidated.log(logStatus.INFO,stepName);
}
if(result.toUpperCase()=="PASS"){
//TextUtils.writeLog(stepName);

if(DStat){ 
Dtest.log(logStatus.PASS,stepName);
}
test.log(logStatus.PASS,stepName);
testConsolidated.log(logStatus.PASS,stepName);
}
if(result.toUpperCase()=="FAIL"){
TestRunner.JiraStat = false;

if(DStat){ 
sFolder = Dfile_path+"\\Screenshots\\"; 
}else{
sFolder = file_path+"\\Screenshots\\";
}


if(DStat){ 
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
}else{
  
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}

}
var img = sFolder+"Image_"+ig+".png";
Sys.Desktop.Picture().SaveToFile(img);
img = "\\Screenshots\\"+"Image_"+ig+".png";
ig++;

if(DStat){ 
Dtest.log(logStatus.FAIL,stepName+test.addScreenCapture("."+img));
test.log(logStatus.FAIL,stepName+test.addScreenCapture(".\\"+file_name+img));
testConsolidated.log(logStatus.FAIL,stepName+testConsolidated.addScreenCapture("..\\"+file_name+img));
}else{
test.log(logStatus.FAIL,stepName+test.addScreenCapture("."+img));
testConsolidated.log(logStatus.FAIL,stepName+testConsolidated.addScreenCapture(".\\"+file_name+img));
}

TextUtils.writeLog(unitName+" is FAILED "+stepName);
//ReportUtils.report.endTest(test);
//ReportUtils.report.flush();
//fileList = slPacker.GetFileListFromFolder(TestRunner.workDir);
//TestRunner.archivePath = TestRunner.packedResults + TestRunner.reportName;
//aqUtils.Delay(4000, "Compressing the Document");
//// Packes the resutls
//if (slPacker.Pack(fileList, TestRunner.workDir, TestRunner.archivePath))
//  Log.Message("Files compressed successfully.");
//  Log.Error(stepName)
//Runner.CallMethod("JIRA.JIRAUpdate");
//ReportUtils.reportConsolidated.endTest(testConsolidated);
//ReportUtils.reportConsolidated.flush();

//var err;
//throw err;
}
if(result.toUpperCase()=="WARNING"){
if(DStat){ 
Dtest.log(logStatus.WARNING,stepName);
}
test.log(logStatus.WARNING,stepName);
testConsolidated.log(logStatus.WARNING,stepName);
}
  
}

function logStep_Screenshot(stepName)
{
if(DStat){ 
 sFolder = Dfile_path+"\\Screenshots\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
} 
}
else{
sFolder = file_path+"\\Screenshots\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}

}
var img = sFolder+"Image_"+ig+".png";
Sys.Desktop.Picture().SaveToFile(img);

img = "\\Screenshots\\"+"Image_"+ig+".png";
ig++;
if((stepName!="")&&(stepName!=null)){
  
if(DStat){ 
Dtest.log(logStatus.INFO,stepName+test.addScreenCapture("."+img));
test.log(logStatus.INFO,stepName+test.addScreenCapture(".\\"+file_name+img));
testConsolidated.log(logStatus.INFO,stepName+testConsolidated.addScreenCapture("..\\"+file_name+img));
}else{
test.log(logStatus.INFO,stepName+test.addScreenCapture("."+img));
testConsolidated.log(logStatus.INFO,stepName+testConsolidated.addScreenCapture(".\\"+file_name+img));
}
}
else
{ 
  
if(DStat){ 
Dtest.log(logStatus.INFO,test.addScreenCapture("."+img));
test.log(logStatus.INFO,test.addScreenCapture(".\\"+file_name+img));
testConsolidated.log(logStatus.INFO,testConsolidated.addScreenCapture("..\\"+file_name+img));
}else{
test.log(logStatus.INFO,test.addScreenCapture("."+img));
testConsolidated.log(logStatus.INFO,testConsolidated.addScreenCapture(".\\"+file_name+img));
} 
 
}
}


function logStep_addImage(stepName)
{


img = "\\Screenshots\\"+"Image_"+ig+".png";
ig++;
if((stepName!="")&&(stepName!=null)){
test.log(logStatus.INFO,stepName+test.addScreenCapture("."+img));
testConsolidated.log(logStatus.INFO,stepName+testConsolidated.addScreenCapture(".\\"+file_name+img));
}
else
{ 
test.log(logStatus.INFO,test.addScreenCapture("."+img));
testConsolidated.log(logStatus.INFO,testConsolidated.addScreenCapture(".\\"+file_name+img));
  
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
