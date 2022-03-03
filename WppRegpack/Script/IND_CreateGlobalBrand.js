//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT CreateClient
function indiaSpecific(State,GST,PAN,TAN,TIN){ 
//Strating Of TestCase

var SaveStat = false;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
//  var indiaspec = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.IndiaSpecific;
  var indiaspec = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl
  //Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 8);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000, "Selecting India");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
Sys.HighlightObject(indiaspec);
var Start = StartwaitTime();
var waitTime = true;
var Difference = 0;
while(waitTime)
if(Difference<61){
if((indiaspec.isEnabled())&&(indiaspec.text=="India Specific")){
Sys.HighlightObject(indiaspec);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
indiaspec.HoverMouse();
indiaspec.Click();
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
 

  var StateCode = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget
  var debtorType = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget
  var C_pan = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
  var C_tin = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.Composite.McTextWidget
  //  var CIN = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
  var C_tan = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget;
  var save = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 
//  var StateCode = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
//  var debtorType = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
//  var C_pan = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.PAN;
//  var C_tan = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.TAN;
//  var C_tin = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);  
  if((StateCode.getText()=="")&&(State!="")){
  Sys.HighlightObject(StateCode);
  StateCode.HoverMouse();
  StateCode.Click();
  DropDownList(State,"State Code")
  SaveStat = true;
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  if((debtorType.getText()=="")&&(GST!="")){
  Sys.HighlightObject(debtorType);
  debtorType.HoverMouse();
  debtorType.Click();
  WorkspaceUtils.DropDownList(GST,"GST Debtor Type")
  SaveStat = true;
  }
  
  if((C_pan.getText()=="")&&(PAN!="")){
  Sys.HighlightObject(C_pan);
  C_pan.HoverMouse();  
   C_pan.setText(PAN);
   SaveStat = true;
  }
  
  if((C_tan.getText()=="")&&(TAN!="")){
   C_tan.setText(TAN);
   SaveStat = true;
  }
  
  if((C_tin.getText()=="")&&(TIN!="")){
  Sys.HighlightObject(C_tin);
  C_tin.HoverMouse();  
   C_tin.setText(TIN);
   SaveStat = true;
  }
  
  if(SaveStat){
//var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.Save;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();

aqUtils.Delay(8000, Indicator.Text);
    var p = Sys.Process("Maconomy");
    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Global Client - India Specific").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  var Okay = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Global Client - India Specific").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
  Okay.Click();
  
}

}

}

function DropDownList(value,feild){ 
var checkMark = false;
Sys.Process("Maconomy").Refresh();
  var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  var Add_Visible4 = true;
  while(Add_Visible4){
  if(list.isEnabled()){
  Add_Visible4 = false;
      for(var i=0;i<list.getItemCount();i++){ 
        if(list.getItem(i).getText_2(0)!=null){ 
          if(list.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(value)!=-1){ 
            list.Keys("[Enter]");
            aqUtils.Delay(1000, "Waiting to find Object");;
            checkMark = true;
            ValidationUtils.verify(true,true,feild+" is selected in Maconomy");
            break;
          }else{
            list.Keys("[Down]");
          }
          
        }else{ 
        Log.Message("i :"+i);
        Log.Message(list.getItem(i).getText_2(0).OleValue.toString().trim());
          list.Keys("[Down]");
        }
      }
  }
  }
  return checkMark;
}

