//USEUNIT EnvParams
//USEUNIT ReportUtils
var img = 0;

function GeneralEvents_OnLogError(Sender, LogParams)
{
  Log.Message(LogParams.MessageText)
 ReportUtils.logStep("FAIL", LogParams.MessageText);
 //Sys.Desktop.Picture().SaveToFile(Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\", reportName);
  
}

function GeneralEvents_OnLogEvent(Sender, LogParams)
{

}

function GeneralEvents_OnLogCheckpoint(Sender, LogParams)
{

// ReportUtils.logStep("INFO", LogParams.MessageText, LogParams.AdditionalText);
}

function GeneralEvents_OnStartTest(Sender)
{
 // VideoRecorder.Start();
}


//function GeneralEvents_OnStopTest(Sender)
//{
//    var projectName = EnvParams.Pname;
//		var versionName = EnvParams.Vname;
//		var cycleName = EnvParams.Cname;
//		var userName = "muthukumar.m@cognizant.com";
//		var  accessKey = "MDA1MDIyZWQtZmEyMC0zOTc4LWI2ZmEtZDM3MTcxMGU1YzRjIDVjYjc1OTJmOWE4NTc5MTA4OTZmZTc5OSBVU0VSX0RFRkFVTFRfTkFNRQ";
//		var secretKey = "jf9LV-GHNp6MKw35xCTPo43WC0V4bwYC4SdsZC5K-Ho"; 
//    var zephyrBaseUrl ="https://prod-api.zephyr4jiracloud.com/connect";
//    if(Log.ErrCount>0){
//    var status = "Passed";// Passed
//    var client = JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.createConnection(zephyrBaseUrl,accessKey,secretKey,userName);
//    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecSttausOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status) 
//    }else{ 
//    var status = "Failed";// Failed
//    var client = JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.createConnection(zephyrBaseUrl,accessKey,secretKey,userName);
//    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecSttausOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status) 
//    
//    }
//    
//}


function GeneralEvents_OnStopTest(Sender)
{

    var projectName = EnvParams.Pname;
		var versionName = EnvParams.Vname;
		var cycleName = EnvParams.Cname;
    var folderName = TestRunner.folderName;
		var testCaseId = TestRunner.testCaseId; 
    var userName = "muthukumar.m@cognizant.com";
		var  accessKey = "MDA1MDIyZWQtZmEyMC0zOTc4LWI2ZmEtZDM3MTcxMGU1YzRjIDVjYjc1OTJmOWE4NTc5MTA4OTZmZTc5OSBVU0VSX0RFRkFVTFRfTkFNRQ";
		var secretKey = "jf9LV-GHNp6MKw35xCTPo43WC0V4bwYC4SdsZC5K-Ho"; 
    var zephyrBaseUrl ="https://prod-api.zephyr4jiracloud.com/connect";

    var client = JavaClasses.com_thed_zephyr_cloud_rest.ZFJCloudRestClient.restBuilder(zephyrBaseUrl, accessKey, secretKey, userName).build();
    if(Log.ErrCount>0){   
    var status = "Failed";// Passed
  JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey, projectName, versionName, cycleName,folderName, testCaseId, status,zephyrBaseUrl,secretKey, userName);
    }    
//  else{      
//    var status = "passed";// Failed
//  JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey, projectName, versionName, cycleName,folderName, testCaseId, status,zephyrBaseUrl,secretKey, userName);
//  
//    }
}

