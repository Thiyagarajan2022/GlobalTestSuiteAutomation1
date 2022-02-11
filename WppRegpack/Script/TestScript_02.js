﻿//USEUNIT TestScript_01
//USEUNIT ActionUtils

function mk(){ 
  Log.Message( TestScript_01.name);
  Log.Message(TestScript_01.mk());
  Indicator.PushText("waiting for window to open");
  aqUtils.Delay(5000,"waiting for window to open");
  
  Delay(5000);
  var maenu = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 2109 Senior Finance A").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  Sys.HighlightObject(maenu);
  Runner.CallMethod("TestScript_01.mk");
  
  var AddIcon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 2109 Senior Finance A").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
  Sys.HighlightObject(AddIcon);
  
  
  var parentObject = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 2109 Senior Finance A");
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName(parentObject,"SingleToolItemControl","Add Job Budget Line (Ctrl+M)");
  Sys.HighlightObject(ObjectAdd);
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName_and_Index(parentObject,"SingleToolItemControl","4");
  Sys.HighlightObject(ObjectAdd);
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withChildCount(parentObject,"SingleToolItemControl","4","0");
  Sys.HighlightObject(ObjectAdd);
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParent(parentObject,"SingleToolItemControl","4","11");
  Sys.HighlightObject(ObjectAdd);
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName_and_Index_withParentClassName(parentObject,"SingleToolItemControl","4","Composite");
  Sys.HighlightObject(ObjectAdd);
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName_conatinsTooltip(parentObject,"SingleToolItemControl","Add Job Budget Line");
  Sys.HighlightObject(ObjectAdd);
  var ObjectAdd = ActionUtils.getObjectAddress_JavaClasssName_withTabText(parentObject,"SingleToolItemControl","Full Budget");;
  Sys.HighlightObject(ObjectAdd);
  
  ImageRepository.Browser_Reporting.WorkspaceClient.Click();
  
}





//1.Create ProjectSuite, Project, TestScript
//2.ImageRepository
//3.TestedAPPs
//4. Find, FindAll, FindChild, FindAllChild with multiple properties
//5. USEUNITS
//6. Runner.callMethod
//7. How to trigeet function.


