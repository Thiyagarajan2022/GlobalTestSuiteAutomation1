﻿//USEUNIT WorkspaceUtils
//USEUNIT CreateExpenses

function justificationPanel(Ereason,Vname,GSTIN,I_no,I_Date){ 
  var Jpanel = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.TabControl;
  WorkspaceUtils.waitForObj(Jpanel);
  ReportUtils.logStep_Screenshot();
  Jpanel.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

var Jtab = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
  WorkspaceUtils.waitForObj(Jtab);
  Jtab.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(2000, Indicator.Text);;
var Exp_Reason = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.McValuePickerWidget
WorkspaceUtils.waitForObj(Exp_Reason);
if(Ereason!=""){
  Exp_Reason.Click()
Exp_Reason.setText(Ereason);
}
WorkspaceUtils.waitForObj(Exp_Reason);
    Sys.Desktop.KeyDown(0x28);
    Sys.Desktop.KeyUp(0x28);
    aqUtils.Delay(1000, Indicator.Text);;
    
var Vendor = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.McValuePickerWidget
WorkspaceUtils.waitForObj(Vendor);
if(Vname!=""){
  Vendor.Click()
Vendor.setText(Vname);
}
WorkspaceUtils.waitForObj(Vendor);
    Sys.Desktop.KeyDown(0x28);
    Sys.Desktop.KeyUp(0x28);
    aqUtils.Delay(1000, Indicator.Text);;
    
var GST = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(GST);
if(GSTIN!=""){
  GST.Click()
GST.setText(GSTIN);
}
WorkspaceUtils.waitForObj(GST);
    Sys.Desktop.KeyDown(0x28);
    Sys.Desktop.KeyUp(0x28);
    aqUtils.Delay(1000, Indicator.Text);;
    
var InvoiceNo = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.McValuePickerWidget;
WorkspaceUtils.waitForObj(InvoiceNo);
if(I_no!=""){
  InvoiceNo.Click()
InvoiceNo.setText(I_no);
}

if(I_Date!=""){    
WorkspaceUtils.waitForObj(InvoiceNo);
    Sys.Desktop.KeyDown(0x28);
    Sys.Desktop.KeyUp(0x28);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
//var InvoiceDate = Aliases.Maconomy.CreateExpense.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
var InvoiceDate =NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.table.McGrid.McDatePickerWidget
 WorkspaceUtils.waitForObj(InvoiceDate);
InvoiceDate.Click()
InvoiceDate.setText(I_Date);
WorkspaceUtils.waitForObj(InvoiceDate);
Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
aqUtils.Delay(1000, Indicator.Text);;
Sys.Desktop.KeyUp(0x10);
Sys.Desktop.KeyUp(0x09);

}
    Sys.Desktop.KeyDown(0x26);
    Sys.Desktop.KeyUp(0x26);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x26);
    Sys.Desktop.KeyUp(0x26);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x26);
    Sys.Desktop.KeyUp(0x26);
    aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x26);
    Sys.Desktop.KeyUp(0x26);
    aqUtils.Delay(1000, Indicator.Text);;
Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
aqUtils.Delay(1000, Indicator.Text);;
Sys.Desktop.KeyUp(0x10);
Sys.Desktop.KeyUp(0x09);
var save = NameMapping.Sys.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.Composite.save
WorkspaceUtils.waitForObj(save);
ReportUtils.logStep_Screenshot();
save.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
    


}