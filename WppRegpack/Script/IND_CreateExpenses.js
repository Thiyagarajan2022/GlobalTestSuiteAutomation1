//USEUNIT ReportUtils
//USEUNIT ValidationUtils


function IndiaSpecific(Reason,Gstin,InvoiceDate,InvoiceNo,VendorName){
  aqUtils.Delay(7000, Indicator.Text);
  var justification = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
  Sys.HighlightObject(justification);
  justification.HoverMouse();
  justification.Click();

    if(ImageRepository.ImageSet0.Expense.Exists()){
        ImageRepository.ImageSet0.Expense.Click();
    }

  var tab = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(tab);
  var reason = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  var gstin = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  var invoicedate = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McDatePickerWidget;
  var invoicenum = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
  var vendorname = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
   
    if(tab.getItem(0).getText_2(0).OleValue.toString().trim()=="Expense Reason"){
      Sys.Desktop.KeyDown(0x27);
      Sys.Desktop.KeyUp(0x27);
        if(Reason!=""){
          Sys.HighlightObject(reason);
          reason.HoverMouse();
          reason.Click();
          reason.setText(Reason);
        }
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(VendorName!=""){
          Sys.HighlightObject(vendorname);
          vendorname.HoverMouse();
          vendorname.Click();
          vendorname.setText(VendorName);
        }
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
        if(Gstin!=""){
          Sys.HighlightObject(gstin);
          gstin.HoverMouse();
          gstin.Click();
          gstin.setText(Gstin);
        }
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(InvoiceNo!=""){
          Sys.HighlightObject(invoicenum);
          invoicenum.HoverMouse();
          invoicenum.Click();
          invoicenum.setText(InvoiceNo);
        }
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      Sys.Desktop.KeyDown(0x27);
      Sys.Desktop.KeyUp(0x27);
      if(InvoiceDate!=""){
          Sys.HighlightObject(invoicedate);
          invoicedate.HoverMouse();
          invoicedate.Click();
          invoicedate.setText(InvoiceDate);
        }
    }  
    for(var j=2;j>0;j--){
        Sys.Desktop.KeyDown(0xA0);
        Sys.Desktop.KeyDown(0x09);
        Sys.Desktop.KeyUp(0xA0);
        Sys.Desktop.KeyUp(0x09);
      }
      var item = tab.getItemCount()
       for(var i=item;i>0;i--){ 
          Sys.Desktop.KeyDown(0x26)
          Sys.Desktop.KeyUp(0x26)																																																																																																																																																																																																													
       }
    var save = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
    Sys.HighlightObject(save);
    save.HoverMouse();
    save.Click();
    var closeexpense = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl;
    closeexpense.HoverMouse();
    Sys.HighlightObject(closeexpense)
    closeexpense.Click();
   ImageRepository.ImageSet0.Forward.Click();
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

