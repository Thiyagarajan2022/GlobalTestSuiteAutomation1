﻿//USEUNIT WorkspaceUtils
function IND_Specific(Unit_Price,OHSN,IHSN,POS){ 
Unit_Price.Keys("[Tab][Tab][Tab]"); 
var OutwardHSN = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//var OutwardHSN = Aliases.Maconomy.PO_JobActivities.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
   if((OHSN!="")&&(OutwardHSN.getText()!=OHSN)){
  OutwardHSN.Click();
  WorkspaceUtils.SearchByValue(OutwardHSN,"Local Specification 8",OHSN,"Outward HSN");
     }
  OutwardHSN.Keys("[Tab]");

var InwardHSN = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget
//var InwardHSN = Aliases.Maconomy.PO_JobActivities.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  if((IHSN!="")&&(InwardHSN.getText()!=IHSN)){
  InwardHSN.Click();
  WorkspaceUtils.SearchByValue(InwardHSN,"Local Specification 9",IHSN,"Inward HSN");
     }
  InwardHSN.Keys("[Tab]");

  var State = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget  
//var State = Aliases.Maconomy.PO_JobActivities.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  if((POS!="")&&(State.getText()!=POS)){
  State.Click();
  WorkspaceUtils.SearchByValue(State,"Local Specification 10",POS,"POS");
     }
     
} 