//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
//USEUNIT WorkspaceUtils

function UAE_Specific(licenceEndDate,licenceNumber){ 
  
 aqUtils.Delay(2000, "Selecting UAE");
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  var UAE = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.journal;
  waitForObj(UAE);
  while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   aqUtils.Delay(100, "Selecting UAE");   
  }
Sys.HighlightObject(UAE);


UAE.HoverMouse();
UAE.Click();

 while(!ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   aqUtils.Delay(100, "Selecting UAE");   
  }


var Licence_End_Date = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.SWTObject("McDatePickerWidget", "", 2);  
var Licence = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
Licence_End_Date.Click();
Licence_End_Date.setText(licenceEndDate);
  aqUtils.Delay(2000, Indicator.Text);
Licence.Click();
Licence.setText(licenceNumber); 
  

         
aqUtils.Delay(2000, Indicator.Text);
var save =Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();



}