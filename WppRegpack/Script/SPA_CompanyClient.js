//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
function spainSpecific(SII_Tax){ 
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  aqUtils.Delay(2000, "Selecting Spain");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var spainspec = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.JobActivities;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000, "Selecting Spain");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
Sys.HighlightObject(spainspec);
var Start = StartwaitTime();
var waitTime = true;
var Difference = 0;
while(waitTime)
if(Difference<61){
if((spainspec.isEnabled())&&(spainspec.text=="Spain")){
Sys.HighlightObject(spainspec);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
spainspec.HoverMouse();
spainspec.Click();
waitTime = false;
}
else{ 
var End = EndTime();
Difference = End - Start;
}
}
else{
 ValidationUtils.verify(true,false,"Screen is not Responding more than a minute");
}

var SIITaxGroup = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.PaytoVendorFromDate;
  if((SIITaxGroup.getText()=="")&&(SII_Tax!="")){
  Sys.HighlightObject(SIITaxGroup);
  SIITaxGroup.Click();
  WorkspaceUtils.SearchByValue(SIITaxGroup, JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language,"Option").OleValue.toString().trim(),SII_Tax,"SII Tax Group");

         
aqUtils.Delay(2000, Indicator.Text);
var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.RemarksSave;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();
aqUtils.Delay(8000, Indicator.Text);
    var p = Sys.Process("Maconomy");
    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Global Client - Spain").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  var Okay = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Global Client - Spain").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
  Okay.Click();
  
}



         } 
         

}