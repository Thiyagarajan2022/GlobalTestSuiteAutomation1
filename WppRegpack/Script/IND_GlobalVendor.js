//USEUNIT ReportUtils
//USEUNIT ValidationUtils


function IndiaSpecific(Vendortype,StateCode,GSTVendor,TDSAplicable,Section,Method){
  aqUtils.Delay(7000, Indicator.Text);
  var indiaspec = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(indiaspec);
var Start = StartwaitTime();
var waitTime = true;
var Difference = 0;
while(waitTime)
if(Difference<61){
if((indiaspec.isEnabled())&&(indiaspec.text=="India Specific")){
Sys.HighlightObject(indiaspec);
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
  var vendortype = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  var Statecode = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
  var GSTVendorType = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  var PermanentEstablishment = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McPopupPickerWidget;
  var TDSapplicable = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McValuePickerWidget;
  var TDSsection =  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Composite.McValuePickerWidget;
  var method = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Composite2.McValuePickerWidget;
  
  if(Vendortype!=""){
  Sys.HighlightObject(vendortype);
  vendortype.HoverMouse();
  vendortype.Click();
  WorkspaceUtils.SearchByValue(vendortype,"Option",Vendortype,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text);  
  
  if(StateCode!=""){
  Sys.HighlightObject(Statecode);
  Statecode.HoverMouse();
  Statecode.Click();
  DropDownList(StateCode,"State Code")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  if(GSTVendor!=""){
  Sys.HighlightObject(GSTVendorType);
  GSTVendorType.HoverMouse();
  GSTVendorType.Click();
  WorkspaceUtils.SearchByValue(GSTVendorType,"Local Specification 7",GSTVendor,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text); 
  
  PermanentEstablishment.Keys("yes");
  aqUtils.Delay(2000, Indicator.Text);
  
  if(TDSAplicable!=""){
  Sys.HighlightObject(TDSapplicable);
  TDSapplicable.HoverMouse();
  TDSapplicable.Click();
  WorkspaceUtils.SearchByValue(TDSapplicable,"Option",TDSAplicable,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text);  
 
    var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(save);
    save.HoverMouse();
    save.Click();
    aqUtils.Delay(2000, Indicator.Text);
    
  if(Section!=""){
  Sys.HighlightObject(TDSsection);
  TDSsection.HoverMouse();
  TDSsection.Click();
  WorkspaceUtils.SearchByValue(TDSsection,"Local Specification 8",Section,"Name");
  }
  aqUtils.Delay(2000, Indicator.Text);  
  
  if(Method!=""){
  Sys.HighlightObject(method);
  method.HoverMouse();
  method.Click();
  WorkspaceUtils.SearchByValue(method,"Option",Method,"Name");
  }
  save.Click();
  aqUtils.Delay(2000, Indicator.Text); 
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

