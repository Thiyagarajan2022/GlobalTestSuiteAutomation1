//USEUNIT ReportUtils
//USEUNIT TextUtils

function specification(workcode,OHSN){ 
  workcode.Keys("[Tab][Tab][Tab][Tab][Tab]");
  var hsn = Aliases.Maconomy.InvoiceOnAccount.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  hsn.Click();
  WorkspaceUtils.SearchByValue(hsn,"Local Specification 8",OHSN,"Outward HSN");
  hsn.Keys("[Tab]");
  
}