//USEUNIT EnvParams
//USEUNIT TestRunner  
//USEUNIT ReportUtils
    var projectName = "";
		var versionName = "";
		var cycleName = "";
    var folderName = "";
    var testCaseId = "";
    var ErrC = 0;
    
function JIRAUpdate(){

		versionName = TestRunner.releasename;
		cycleName = TestRunner.cyclename;
    folderName = "";
		testCaseId = TestRunner.testCaseId;  
     projectName = EnvParams.Pname;
     	var userName = EnvParams.JiraUsername;
		var  accessKey = EnvParams.JiraAccessKey;
		var secretKey = EnvParams.JiraSecrekey;
    var zephyrBaseUrl =EnvParams.JirazephyrBaseUrl;

    var entityName = "execution";
    var expirationInsec = 360;
    var comment = "TestReport_Uploaded_Successfully_in_JIRA.";
    Log.Message("TestRunner.JiraUpdate :"+TestRunner.JiraUpdate)
    Log.Message("projectName :"+projectName)
    Log.Message("cycleName :"+cycleName)
    Log.Message("folderName :"+folderName)
    Log.Message("testCaseId :"+testCaseId)

    var client = JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.createConnection(zephyrBaseUrl,accessKey,secretKey,userName);
    if(!TestRunner.JiraStat){
    var status = "Failed";// Passed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
//        if(ReportUtils.DStat){ 
//    folderName = TestRunner.folderName;
//var workDir = ReportUtils.Dfile_path+"\\";
//Log.Message(workDir)
//var fileList = slPacker.GetFileListFromFolder(workDir);
//var archivePath = ReportUtils.file_path +"\\"+ ReportUtils.Dfile_name;
//Delay(5000);
//if (slPacker.Pack(fileList, workDir, archivePath))
//  Log.Message("Files compressed successfully");
//Delay(4000);
////     JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,archivePath+".zip",expirationInsec, comment) 
//    }else{
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,TestRunner.archivePath+".zip",expirationInsec, comment)
//    }
    }     
  else{      
    var status = "passed";// Failed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
//    if(ReportUtils.DStat){ 
//    folderName = TestRunner.folderName;
//      Log.Message("JIRA Attachment :"+ReportUtils.Dfile_path+"\\"+EnvParams.Opco+"_"+TestRunner.unitName+"_TestLog.txt")
//     JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,ReportUtils.Dfile_path+"\\"+EnvParams.Opco+"_"+TestRunner.unitName+"_TestLog.txt",expirationInsec, comment) 
//    }else{
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,ReportUtils.file_path+"\\"+EnvParams.Opco+"_"+TestRunner.unitName+"_TestLog.txt",expirationInsec, comment)
//    }
    }
    
    TestRunner.JiraUpdate = false;
    Log.Message("TestRunner.JiraUpdate :"+TestRunner.JiraUpdate)
}

function JIRAErrorUpdate(){

		versionName = TestRunner.releasename;
		cycleName = TestRunner.cyclename;
    folderName = TestRunner.folderName;
		testCaseId = TestRunner.testCaseId;  
     projectName = EnvParams.Pname;
     	var userName = EnvParams.JiraUsername;
		var  accessKey = EnvParams.JiraAccessKey;
		var secretKey = EnvParams.JiraSecrekey;
    var zephyrBaseUrl =EnvParams.JirazephyrBaseUrl;
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
    var status = "Failed";// Passed
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.UpdateExecStatusOfTestCase(client,accessKey,projectName,versionName,cycleName,folderName,testCaseId,status)
    JavaClasses.com_cts_ZephyrApiUsecases.UpdateExecutionStatus.addAttachements(client,accessKey,entityName,TestRunner.archivePath+".zip",expirationInsec, comment)

  

}

