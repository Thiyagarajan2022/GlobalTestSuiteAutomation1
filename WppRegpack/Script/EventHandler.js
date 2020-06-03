﻿//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT TestRunner 
var img = 0;

function GeneralEvents_OnLogError(Sender, LogParams)
{
  try{
  Log.Message(LogParams.MessageText)
 ReportUtils.logStep("FAIL", LogParams.MessageText);
 }catch(e){ 
   
 }
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
   // var userName = "muthukumar.m@cognizant.com";
//		var  accessKey = "MDA1MDIyZWQtZmEyMC0zOTc4LWI2ZmEtZDM3MTcxMGU1YzRjIDVjYjc1OTJmOWE4NTc5MTA4OTZmZTc5OSBVU0VSX0RFRkFVTFRfTkFNRQ";
//		var secretKey = "jf9LV-GHNp6MKw35xCTPo43WC0V4bwYC4SdsZC5K-Ho"; 
  //  var zephyrBaseUrl ="https://prod-api.zephyr4jiracloud.com/connect";
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
