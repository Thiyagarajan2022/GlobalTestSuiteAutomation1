//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT TestRunner 
var img = 0;

function GeneralEvents_OnLogError(Sender, LogParams)
{
  Log.Message(LogParams.MessageText)
 ReportUtils.logStep("FAIL", LogParams.MessageText);
}

function GeneralEvents_OnLogEvent(Sender, LogParams)
{

}

function GeneralEvents_OnLogCheckpoint(Sender, LogParams)
{

}

function GeneralEvents_OnStartTest(Sender)
{

}


function GeneralEvents_OnStopTest(Sender)
{

  var automationStat_Date = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%Y");
  var filePath = TestRunner.exeResults+"RunTime_statistics_"+automationStat_Date+".xlsx";
  var exeTime ="";
  // Create RuntimeStatistics file if not exists. (If first testcase itself failed this steps will be executed)
  if(!aqFile.Exists(filePath))
      ExcelUtils.create_AutomationStat_Excel(filePath);  
      
   // Calculate Runtime value and publish to excel.   
   eTime = new Date(); 
   exeTime = WorkspaceUtils.timeDifference(TestRunner.sTime, eTime); 

   if(!TestRunner.testCase_Stat_updated_flag)  
       ExcelUtils.writeTo_AutomationStat_Excel(filePath,TestRunner.testName,exeTime);
   ExcelUtils.close_AutomationStat_Excel(); 
   TextUtils.writeLog(TestRunner.testName +" Execution End Time :"+eTime); 
  
    var projectName = EnvParams.Pname;
		var versionName = TestRunner.releasename;
		var cycleName = TestRunner.cyclename;
    var folderName = TestRunner.folderName;
		var testCaseId = TestRunner.testCaseId;  
    var userName = EnvParams.JiraUsername;
		var  accessKey = EnvParams.JiraAccessKey;
		var secretKey = EnvParams.JiraSecrekey;
    var zephyrBaseUrl =EnvParams.JirazephyrBaseUrl;   
    var entityName = "execution";
    var expirationInsec = 360;
    var comment = "TestReport_Uploaded_Successfully_in_JIRA.";
    if((folderName!="")||(folderName!=null)){ 
      
    }else{ 
      folderName = "";
    }
    
    Log.Message(projectName)
    Log.Message(versionName)
    Log.Message(cycleName)
    Log.Message(folderName)
    Log.Message(testCaseId)
    Log.Message("TestRunner.JiraUpdate :"+TestRunner.JiraUpdate)
if(TestRunner.JiraUpdate){
    var client = JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.createConnection(zephyrBaseUrl,accessKey,secretKey,userName);
    if(!TestRunner.JiraStat){ 
//    var status = "Failed";// Passed
//    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
//    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,TestRunner.archivePath+".zip",expirationInsec, comment)

    var status = "Failed";// Passed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
        if(ReportUtils.DStat){ 
var workDir = ReportUtils.Dfile_path+"\\";
Log.Message(workDir)
var fileList = slPacker.GetFileListFromFolder(workDir);
var archivePath = ReportUtils.file_path +"\\"+ ReportUtils.Dfile_name;
Log.Message(archivePath)
Delay(5000);
if (slPacker.Pack(fileList, workDir, archivePath))
  Log.Message("Files compressed successfully");
Delay(4000);
     JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,archivePath+".zip",expirationInsec, comment) 
    }else{
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,TestRunner.archivePath+".zip",expirationInsec, comment)
    }
    }   
    } 

}
