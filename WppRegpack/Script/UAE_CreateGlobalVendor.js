//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
//USEUNIT WorkspaceUtils

function UAE_Specific(licenceEndDate,licenceNumber){ 
  
 aqUtils.Delay(2000, "Selecting UAE");
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var specificCountry = ""
  if(EnvParams.Country.toUpperCase()=="UAE")     
  specificCountry = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.journal;
  if(EnvParams.Country.toUpperCase()=="QATAR")
  specificCountry = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2
  if(EnvParams.Country.toUpperCase()=="EGYPT")
  specificCountry = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8)
  waitForObj(specificCountry);
  while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   aqUtils.Delay(100, "Selecting UAE");   
  }
Sys.HighlightObject(specificCountry);


specificCountry.HoverMouse();
specificCountry.Click();

 while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   aqUtils.Delay(100, "Selecting UAE");   
  }

var Licence_End_Date =""
var Licence = ""
if(EnvParams.Country.toUpperCase()=="UAE")   
{
  Licence_End_Date= Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.SWTObject("McDatePickerWidget", "", 2);  
  Licence = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
}
if(EnvParams.Country.toUpperCase()=="QATAR")
{
  Licence_End_Date= Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2)
  Licence = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2)
}  

if(EnvParams.Country.toUpperCase()=="EGYPT")   
{
 Licence_End_Date= Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 7).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2)
 Licence = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.SWTObject("Composite", "", 7).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2)
}

Licence_End_Date.Click();

Licence_End_Date.setText(licenceEndDate);
  aqUtils.Delay(2000, Indicator.Text);
Licence.Click();
Licence.setText(licenceNumber); 
  

         
aqUtils.Delay(2000, Indicator.Text);
var save =
Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();



}