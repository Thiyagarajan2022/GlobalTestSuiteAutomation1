﻿//USEUNIT ReportUtils

function GeneralEvents_OnLogError(Sender, LogParams)
{
  Log.Message(LogParams.MessageText)
 ReportUtils.logStep("FAIL", LogParams.MessageText);
 //Sys.Desktop.Picture().SaveToFile(Project.Path+TextUtils.GetProjectValue("ReportPath")+"Report_"+ReportDate+"\\", reportName);
  
}

function GeneralEvents_OnLogEvent(Sender, LogParams)
{
//ReportUtils.logStep("INFO", LogParams.MessageText, "");
//ReportUtils.logStep("PASS", LogParams.MessageText, "");
}

function GeneralEvents_OnLogCheckpoint(Sender, LogParams)
{

 ReportUtils.logStep("INFO", LogParams.MessageText, LogParams.AdditionalText);
}

function GeneralEvents_OnStartTest(Sender)
{
 // VideoRecorder.Start();
}


function GeneralEvents_OnStopTest(Sender)
{
 // VideoRecorder.Stop()
}