//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
//USEUNIT WorkspaceUtils

function UAE_Specific(Licence_No,Licence_EndDate){ 
  
 aqUtils.Delay(2000, "Selecting UAE");
//Strating Of TestCase
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  var UAE = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.journal;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000, "Selecting UAE");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
Sys.HighlightObject(UAE);


Sys.HighlightObject(UAE);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
UAE.HoverMouse();
UAE.Click();
waitTime = false;
  aqUtils.Delay(2000, "Selecting UAE");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000, "Selecting UAE");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }


var Licence_End_Date = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2);
var Licence = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2)
Licence_End_Date.Click();
Licence_End_Date.setText(Licence_EndDate);
  aqUtils.Delay(2000, Indicator.Text);
Licence.Click();
Licence.setText(Licence_No); 
  

         
aqUtils.Delay(2000, Indicator.Text);
var save = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.newassetbutton;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();

aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Country Client - UAE").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Country Client - UAE").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
}


}