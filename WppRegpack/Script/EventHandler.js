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
    var client = JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.createConnection(zephyrBaseUrl,accessKey,secretKey,userName);
    if(Log.ErrCount>0){   
    var status = "Failed";// Passed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,TestRunner.archivePath+".zip",expirationInsec, comment)

    }    

}
