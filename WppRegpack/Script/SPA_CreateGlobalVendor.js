//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateGlobalVendor
function spainSpecific(SII_Tax){ 
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
} 
  var spainspec = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
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

var SIITaxGroup = Aliases.Maconomy.BatchAccrual.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
  if(SII_Tax!=""){
  Sys.HighlightObject(SIITaxGroup);
  SIITaxGroup.Click();
  WorkspaceUtils.SearchByValue(SIITaxGroup, JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,CreateGlobalVendor.Language,"Option").OleValue.toString().trim(),SII_Tax,"SII Tax Group");
         } 
         
  aqUtils.Delay(2000, Indicator.Text);
var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();
}