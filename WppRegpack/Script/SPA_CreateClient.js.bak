//USEUNIT ReportUtils
//USEUNIT ValidationUtils
function spainSpecific(SII_Tax){ 
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 
ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
} 
  var spainspec = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.IndiaSpecific;
Sys.HighlightObject(spainspec);
var Start = StartwaitTime();
var waitTime = true;
var Difference = 0;
while(waitTime)
if(Difference<61){
if((spainspec.isEnabled())&&(spainspec.text=="Spain")){
Sys.HighlightObject(spainspec);
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

var SIITaxGroup = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  if(SII_Tax!=""){
  Sys.HighlightObject(SIITaxGroup);
  SIITaxGroup.Click();
  WorkspaceUtils.SearchByValue(SIITaxGroup,"Option",SII_Tax,"SII Tax Group");
         } 
         
  aqUtils.Delay(2000, Indicator.Text);
var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.Save;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();
}