﻿//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
function spainSpecific(SII_Tax){ 
//Strating Of TestCase
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
  
  var spainspec = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 9)          
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(2000, "Selecting Spain");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
   waitForObj(spainspec);
Sys.HighlightObject(spainspec);

Sys.HighlightObject(spainspec);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

spainspec.HoverMouse();
spainspec.Click();
spainspec.Click();


var SIITaxGroup = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  if(SII_Tax!=""){
  Sys.HighlightObject(SIITaxGroup);
  SIITaxGroup.Click();
  WorkspaceUtils.SearchByValue(SIITaxGroup, JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,CreateClient.Language,"Option").OleValue.toString().trim(),SII_Tax,"SII Tax Group");
         } 
         
  aqUtils.Delay(2000, Indicator.Text);
var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.Save;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();
}