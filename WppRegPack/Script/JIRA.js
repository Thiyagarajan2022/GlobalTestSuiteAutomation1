﻿//USEUNIT EnvParams
//USEUNIT TestRunner  
    var projectName = "";
		var versionName = "";
		var cycleName = "";
    var folderName = "";
    var testCaseId = "";
function JIRAUpdate(folderName,testCaseId){

     projectName = EnvParams.Pname;
		 versionName = EnvParams.Vname;
		 cycleName = EnvParams.Cname;
//     folderName = "1201-NewFolder";
//		 testCaseId = "TSTAUTO-1"; 
//Log.Message(projectName +" "+versionName+" "+cycleName+" "+folderName+" "+testCaseId)
		var userName = "muthukumar.m@cognizant.com";
		var  accessKey = "MDA1MDIyZWQtZmEyMC0zOTc4LWI2ZmEtZDM3MTcxMGU1YzRjIDVjYjc1OTJmOWE4NTc5MTA4OTZmZTc5OSBVU0VSX0RFRkFVTFRfTkFNRQ";
		var secretKey = "jf9LV-GHNp6MKw35xCTPo43WC0V4bwYC4SdsZC5K-Ho"; 
    var zephyrBaseUrl ="https://prod-api.zephyr4jiracloud.com/connect";

    var client = JavaClasses.com_thed_zephyr_cloud_rest.ZFJCloudRestClient.restBuilder(zephyrBaseUrl, accessKey, secretKey, userName).build();
    if(Log.ErrCount>0){   
    var status = "Failed";// Passed
  JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey, projectName, versionName, cycleName,folderName, testCaseId, status,zephyrBaseUrl,secretKey, userName);
  JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client," TestReport Attached  in JIRA",TestRunner.archivePath);
    }    
  else{      
    var status = "passed";// Failed
  JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey, projectName, versionName, cycleName,folderName, testCaseId, status,zephyrBaseUrl,secretKey, userName);
  JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client," TestReport Attached  in JIRA",TestRunner.archivePath);
    }
}

