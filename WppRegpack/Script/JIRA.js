﻿//USEUNIT EnvParams
//USEUNIT TestRunner  
//USEUNIT ReportUtils
    var projectName = "";
		var versionName = "";
		var cycleName = "";
    var folderName = "";
    var testCaseId = "";
function JIRAUpdate(folderName,testCaseId,versionName,cycleName){

   projectName = EnvParams.Pname;
   var userName = EnvParams.JiraUsername;
		var  accessKey = EnvParams.JiraAccessKey;
		var secretKey = EnvParams.JiraSecrekey;
    var zephyrBaseUrl =EnvParams.JirazephyrBaseUrl;
    
    Log.Message(projectName)
    Log.Message(versionName)
    Log.Message(cycleName)
    Log.Message(folderName)
    Log.Message(testCaseId)
      Log.Message(accessKey)
        Log.Message(secretKey)
         Log.Message(zephyrBaseUrl)
//		 versionName = EnvParams.Vname;
//		 cycleName = EnvParams.Cname;
//     cycleName = EnvParams.Country;
//		var userName = "muthukumar.m@cognizant.com";
//		var  accessKey = "MDA1MDIyZWQtZmEyMC0zOTc4LWI2ZmEtZDM3MTcxMGU1YzRjIDVjYjc1OTJmOWE4NTc5MTA4OTZmZTc5OSBVU0VSX0RFRkFVTFRfTkFNRQ";
//		var secretKey = "jf9LV-GHNp6MKw35xCTPo43WC0V4bwYC4SdsZC5K-Ho"; 
//    var zephyrBaseUrl ="https://prod-api.zephyr4jiracloud.com/connect";
    var entityName = "execution";
    var expirationInsec = 360;
    var comment = "TestReport_Uploaded_Successfully_in_JIRA.";

    var client = JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.createConnection(zephyrBaseUrl,accessKey,secretKey,userName);
    if(Log.ErrCount>0){   
    var status = "Failed";// Passed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,TestRunner.archivePath+".zip",expirationInsec, comment)

    }     
  else{      
    var status = "passed";// Failed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,ReportUtils.file_path+"\\"+EnvParams.Opco+"_"+TestRunner.unitName+"_TestLog.txt",expirationInsec, comment)
    }
}

