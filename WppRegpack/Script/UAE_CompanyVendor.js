//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
//USEUNIT WorkspaceUtils

function UAE_Specific(licenceEndDate,licenceNumber){ 
  
 aqUtils.Delay(2000, "Selecting UAE");
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  var UAE = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.JobActivities;
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


var Licence_End_Date = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2);
var Licence = Aliases.Maconomy.ChangePaymentSelection.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.SWTObject("McTextWidget", "", 2);
Licence_End_Date.Click();
Licence_End_Date.setText(licenceEndDate);
  aqUtils.Delay(2000, Indicator.Text);
Licence.Click();
Licence.setText(licenceNumber); 
  

         
aqUtils.Delay(2000, Indicator.Text);
var save =Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.RemarksSave;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();



}