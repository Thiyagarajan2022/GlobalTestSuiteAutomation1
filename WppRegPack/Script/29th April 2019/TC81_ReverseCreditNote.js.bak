//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "ReverseCreditNote";
var invoiceDetails = [];
var arrays = [];
var arrys = [];
var addArray = [];
var removed = [];
var samearr = []; 
var changed = false;

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountPayable.Exists()){
ImageRepository.ImageSet.AccountPayable.Click();// GL
}
else if(ImageRepository.ImageSet.AccountPayable2.Exists()){
ImageRepository.ImageSet.AccountPayable2.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
}

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|AP Transactions");
}

}
Delay(6000);
}


function invoiceAllocation(){ 


  var invoiceAllocations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  invoiceAllocations.Click();
  Delay(4000);
  if(ImageRepository.ImageSet.Close_Filter.Exists()){ 
  ImageRepository.ImageSet.Close_Filter.Click();
  Delay(3000);
}
  var newInvoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 5);
  newInvoice.Click();
  Delay(4000);
  invoiceDetails = SOXexcel(sheetName,1);
  var company = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(invoiceDetails[0]!=""){
  company.Click();
  WorkspaceUtils.SearchByValue(company,"Company",invoiceDetails[0]);
    }else{ 
    ValidationUtils.verify(false,true,"Company Number is Need to create VendorInvoice");
  }
  
  var transaction = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
  if(invoiceDetails[1]!=""){
  transaction.Click();
  var ObjectAddrs = transaction;
  var popupName = "Transaction Type";
  var value = invoiceDetails[1];
  var checkmark = false;
  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    code.Keys("[Tab]");
    Delay(3000);
    
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(invoiceDetails[0]);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if((table.getItem(i).getText_2(0).OleValue.toString().trim()==value)&&(table.getItem(i).getText_2(1).OleValue.toString().trim()==invoiceDetails[0])){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          checkmark = true;
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
      ObjectAddrs.setText("");
    }
    }else{ 
    ValidationUtils.verify(false,true,"Transaction Type is Need to create VendorInvoice");
  }
    
  var EntryDate = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2).getText();
  
  var InvoiceNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  InvoiceNo.setText(invoiceDetails[2])
  
  var InvoiceDate = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McDatePickerWidget", "", 2);
  InvoiceDate.setText(EntryDate);
  Delay(2000);
  
  
  var Description = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
  Description.setText(invoiceDetails[3])
  Delay(2000);
  
  var screens = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");
  screens.Click();
  screens.MouseWheel(-1);
  
    
  var VendorNum = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(invoiceDetails[5]!=""){
    VendorNum.Click();
  var ObjectAddrs = VendorNum;
  var popupName = "Vendor";
  var value = invoiceDetails[4];
  var checkmark = false;
  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)

 var AllVendor = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Vendors");
 AllVendor.Click();
 Delay(4000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
//    code.Keys("[Tab]");
    Delay(3000);

    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          checkmark = true;
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
      ObjectAddrs.setText("");
    }
    }
  
 var PaymentMode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
 if(invoiceDetails[5]!=""){
  PaymentMode.Click();
  WorkspaceUtils.SearchByValue(PaymentMode,"Payment Mode",invoiceDetails[5]);
    }else{ 
    ValidationUtils.verify(false,true,"Payment Mode is Need to create VendorInvoice");
  }
 var journalNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McValuePickerWidget", "", 2);
  if(invoiceDetails[6]!=""){
  journalNo.Click();
  var ObjectAddrs = journalNo;
  var popupName = "Vendor Invoice";
  var value = invoiceDetails[6];
//  WorkspaceUtils.SearchByValue(journalNo,"Vendor Invoice",invoiceDetails[6]);
  var checkmark = false;
  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Invoice").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          checkmark = true;
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
      ObjectAddrs.setText("");
    }


    }else{ 
    ValidationUtils.verify(false,true,"JournalNo is Need to create VendorInvoice");
  }
 
  
  var reverceCopy = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  if(!reverceCopy.getSelection()){ 
  reverceCopy.Click();
  }
  else
{ 
    reverceCopy.Click();
}
 Delay(2000);
   var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
  Sys.HighlightObject(save);
  save.Click();
  Delay(5000);
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="AP Transactions - Invoice Allocation"){
var OK = Sys.Process("Maconomy").SWTObject("Shell", "AP Transactions - Invoice Allocation").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
OK.Click();

}
Delay(5000);
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------

var invoiceAllocationLine = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - SSC IN -  Senior AP").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
invoiceAllocationLine.Click();
Delay(1000);

table  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
// Getting Details from Invoice Line
    var a = 0;
    var b=0;
       removed = [];
       addArray = [];
    arrys = excel(sheetName,2)
       for(var i=0;i<arrys.length;i++){
       var sp = arrys[i].split("*");
       if(sp[sp.length-2].toUpperCase().indexOf("CHANGE")!=-1){
       vv = arrys[i].lastIndexOf("*");
       subStr = arrys[i].substring(0,vv);
       vv = subStr.lastIndexOf("*");
         addArray[a] = subStr.substring(0,vv+1);
         a++;
       }



       if(sp[sp.length-2].toUpperCase().indexOf("REMOVE")!=-1){ 
       vv = arrys[i].lastIndexOf("*");
       subStr = arrys[i].substring(0,vv);
       vv = subStr.lastIndexOf("*");
         removed[b] = subStr.substring(0,vv+1);
         Log.Message("Remove :"+removed[b]);
         Log.Message("b"+b);
         b++;
       }
       Log.Message("removed Length :"+removed.length);
       }  
    
    for(var j=0;j<table.getItemCount()-1;j++){  
      var temp = "";
      for(var i=0;i<table.getColumnCount();i++){ 
      if((i==1)||(i==3)||(i==4)||(i==5)||(i==7)||(i==8)||(i==13)||(i==15)||(i==29)){
            if(table.getItem(j).getText_2(i)!=""){
       temp = temp+ table.getItem(j).getText_2(i).OleValue.toString().trim()+"*";
       }
       else{ 
         temp= temp+"*";
       }
       }
      }
      arrays [j] = temp;
      Log.Message("Maconomy :"+temp);
      }
      


// Finding given romoved item is present in maconomy
 var z=0;
 var y=0;
 var x=0;
 var temp;
    for(var i=0;i<removed.length;i++){ 
    temp = false;
      for(var j=0;j<arrays.length;j++){ 
         if(removed[i]==arrays[j]){ 
           samearr[z]=arrays[j];   //matched array with duplicate
           Log.Message("SAME :"+samearr[z]);
           z++;
           temp = true;
         }
         
    }  
    }
             
         for(var i=0;i<samearr.length;i++){ 
           Log.Message("Removed :"+samearr[i])
         }

         
         
         if(removed.length==samearr.length){ 
         
         }else{ 
           Log.Warning("Given purchase order details to remove from invoice is not found");
         }

//Remove from Invoice Line

var invoiceAllocationLine = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - SSC IN -  Senior AP").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
Sys.HighlightObject(invoiceAllocationLine)  ;
invoiceAllocationLine.Click();
Delay(1000);

budgetTable  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
  
//    var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
//    Sys.HighlightObject(FullBudget)  ;
//    FullBudget.Click();
//    var budgetTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
//    Log.Message("Budget Table :"+budgetTable.getItemCount());
    var newtable = budgetTable.getItemCount()
    if(newtable>1){ 
    Sys.Desktop.KeyDown(0x09)
    Delay(2000);
    Sys.Desktop.KeyDown(0x09)
    Delay(2000);
    Sys.Desktop.KeyDown(0x09)
    Delay(2000);
//    Sys.Desktop.KeyDown(0x09)
//    Delay(2000);
//    Sys.Desktop.KeyDown(0x09)
//    Delay(2000);
//    Sys.Desktop.KeyDown(0x09)
//    Delay(2000);
//    Sys.Desktop.KeyUp(0x09)
//    Sys.Desktop.KeyUp(0x09)
//    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    }
    var Jtemp = true;
    for(var j=0;j<newtable-1;j++){
////    Log.Message("newtable-  :"+newtable);
//    Log.Message("J :"+j)
    if(!Jtemp){ 
      j=0;
    }

    var workcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2)
     var temp = "";
      for(var i=0;i<budgetTable.getColumnCount();i++){ 
       if((i==1)||(i==3)||(i==4)||(i==5)||(i==7)||(i==8)||(i==13)||(i==15)||(i==29)){
      if(budgetTable.getItem(j).getText_2(i)!=""){
       temp = temp+ budgetTable.getItem(j).getText_2(i).OleValue.toString().trim()+"*";
       }
       else{ 
         temp= temp+"*";
       }
      }
      }
      var wc= true;
      Jtemp = true;
      for(var x=0;x<removed.length;x++){
      
      if(temp==removed[x]){ 
//      Log.Message("x :"+removed.length);
//      Log.Message("x value :"+x);
//      Log.Message("temp :"+temp);
//      Log.Message("removed :"+removed[x]);
      removed[x]="";
        var del_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6)
      Sys.HighlightObject(del_budget);
      del_budget.Click();
      var sub_delete = Sys.Process("Maconomy").SWTObject("Shell", "Delete").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      sub_delete.Click();
      Delay(4000);
      changed = true;
      Jtemp = false;
      if(j!=0){
      j=j-1;
      
      }
      
      newtable=budgetTable.getItemCount();
      wc=false;
//      del_budget.Click();
//      var sub_delete = Sys.Process("Maconomy").SWTObject("Shell", "Delete").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//      sub_delete.Click();
        
      Delay(4000);
      break;
      }
    }
//    Log.Message(temp);
    newtable = budgetTable.getItemCount();
    Log.Message("Re-Count :"+(newtable-2))
    Log.Message("j :"+j);
    if(wc){
    if(j!=(newtable-2)){
    workcode.Keys("[Down]");
    }
//    newtable = newtable-1;
//    Log.Message("newtable:"+newtable);
    }
    }
    var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3)
    Sys.HighlightObject(Save);
    if(Save.isEnabled()){
    Delay(3000);
    Save.Click();
    }

    
    budgetTable  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    var newtable = budgetTable.getItemCount()
        for(var j=0;j<newtable;j++){
        budgetTable.Keys("[Up]");
        
        }
// addArray
      for(var j=0;j<newtable-1;j++){
var invoiceAllocationLine = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - SSC IN -  Senior AP").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
Sys.HighlightObject(invoiceAllocationLine)  ;
invoiceAllocationLine.Click();
Delay(1000);
      for(k=0;k<addArray.length;k++){
      var newItem = addArray[k].split("*");
      if((budgetTable.getItem(j).getText_2(1).OleValue.toString().trim()==newItem[0]) && (budgetTable.getItem(j).getText_2(3).OleValue.toString().trim()==newItem[1]) &&
      (budgetTable.getItem(j).getText_2(4).OleValue.toString().trim()==newItem[2])){ 
//      for(var i=0;i<budgetTable.getColumnCount();i++){ 
    Sys.Desktop.KeyDown(0x09)
    Delay(2000);
    Sys.Desktop.KeyDown(0x09)
    Delay(2000);
    Sys.Desktop.KeyDown(0x09)
    Delay(2000);
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09) 
    Delay(2000)  
    
    workcodeDescription = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    if((newItem[3]!="") && (budgetTable.getItem(j).getText_2(5).OleValue.toString().trim()!=newItem[3])){
    workcodeDescription.Click();
    workcodeDescription.setText(newItem[3])
    }
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Sys.Desktop.KeyUp(0x09)
    Quantity = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    if((newItem[4]!="") && (budgetTable.getItem(j).getText_2(7).OleValue.toString().trim()!=newItem[4])){
    Quantity.Click();
    Quantity.setText(newItem[4])
    }
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Sys.Desktop.KeyUp(0x09) 
    UnitPrice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    if((newItem[5]!="") && (budgetTable.getItem(j).getText_2(8).OleValue.toString().trim()!=newItem[5])){
    UnitPrice.Click();
    UnitPrice.setText(newItem[5])
    }
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Delay(2000)
    tax_code_1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    if((newItem[6]!="") && (budgetTable.getItem(j).getText_2(13).OleValue.toString().trim()!=newItem[5])){
    tax_code_1.Click();
    WorkspaceUtils.SearchByValue(tax_code_1,"G/L Tax Code",newItem[6]);
    }
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    tax_code_2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    if((newItem[7]!="") && (budgetTable.getItem(j).getText_2(15).OleValue.toString().trim()!=newItem[7])){
    tax_code_2.Click();
    WorkspaceUtils.SearchByValue(tax_code_2,"G/L Tax Code",newItem[7]);
    }
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
        Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
    Delay(2000)
    Sys.Desktop.KeyDown(0x09)
//    Delay(2000)
    
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
    Sys.Desktop.KeyUp(0x09)
//    Delay(2000)
    HSN = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    if((newItem[8]!="") && (budgetTable.getItem(j).getText_2(29).OleValue.toString().trim()!=newItem[8])){
    HSN.Click();
    WorkspaceUtils.SearchByValue(tax_code_2,"Local Specification 9",newItem[8]);
    }
    for(v=0;v<26;v++){
    Delay(2000);
    Sys.Desktop.KeyDown(0x10)
    Sys.Desktop.KeyDown(0x09)    
    Sys.Desktop.KeyUp(0x10)
    Sys.Desktop.KeyUp(0x09)
    
    }
    
    var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
    save.Click();    
      Delay(5000);
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="AP Transactions - Invoice Allocation"){
var OK = Sys.Process("Maconomy").SWTObject("Shell", "AP Transactions - Invoice Allocation").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
OK.Click();

}            
              
//      }      
//                  }
}
if(j<newtable-2)
budgetTable.Keys("[Down]");
Delay(2000);
}
        
        }
















//-----------------------------------------------------------------------------------------------------------------------------------------------------------------
  
    Delay(2000); 
var invoiceAllocationLine = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - SSC IN -  Senior AP").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
Sys.HighlightObject(invoiceAllocationLine)  ;
invoiceAllocationLine.Click();
Delay(2000);
budgetTable  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    var itemCount = budgetTable.getItemCount();
    var totall=0.00;
    var Tax1 = "";
    var Tax2 = "";
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(budgetTable.getItem(i).getText_2(9).OleValue.toString().trim()!=""){ 
      if(i==0){
      Tax1 = budgetTable.getItem(i).getText_2(13).OleValue.toString().trim();
      Tax2 = budgetTable.getItem(i).getText_2(15).OleValue.toString().trim();
      }
      Log.Message("budgetTable :"+budgetTable.getItem(i).getText_2(9))
      var temp = aqConvert.StrToFloat(budgetTable.getItem(i).getText_2(9));
      totall = totall+temp;
      Log.Message(totall) ;
      }
      }
      }
      
      
var excl_Tax = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(excl_Tax);
excl_Tax.setText(totall);

var tax_code_1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McValuePickerWidget", "", 2);
if(Tax1!=""){ 
   tax_code_1.Click();
  WorkspaceUtils.SearchByValue(tax_code_1,"G/L Tax Code",Tax1);
}
else{ 
  tax_code_1.Click();
  tax_code_1.Keys("^a[BS]")
}
var tax_code_2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McValuePickerWidget", "", 2);
if(Tax2!=""){ 
   tax_code_2.Click();
  WorkspaceUtils.SearchByValue(tax_code_2,"G/L Tax Code",Tax2);
}else{ 
  tax_code_2.Click();
  tax_code_2.Keys("^a[BS]")
}
Delay(2000);
var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4);
save.Click();
  Delay(4000);
var action = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("GroupToolItemControl", "", 9);
action.Click();
  Delay(3000);
  Sys.Process("Maconomy").Refresh();
  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
  Sys.HighlightObject(table);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  Delay(4000);  
var Add_Visible0 = true;
var loading = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 5);
while(Add_Visible0){
if(loading.isEnabled()){
Add_Visible0 = false;
Delay(2000);
}
}


 action.Click();
// action.Click();
  Delay(3000);
  Sys.Process("Maconomy").Refresh();
  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
  Sys.HighlightObject(table);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  Delay(4000);  
var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
dicratory.Keys("C:\\Users\\674087\\Desktop\\New folder\\test1.xlsx");
var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
Sys.HighlightObject(opendoc);
opendoc.Click();
Delay(3000);

action.Click();
// action.Click();
  Delay(3000);
  Sys.Process("Maconomy").Refresh();
  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
  Sys.HighlightObject(table);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x28);
  Delay(1000);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(1000);
  Sys.Desktop.KeyDown(0x0D);
  Sys.Desktop.KeyUp(0x0D);
  Delay(20000); 
 

  if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="AP Transactions - Invoice Allocation"){
  Log.Message(Sys.Process("Maconomy").SWTObject("Shell", "AP Transactions - Invoice Allocation").SWTObject("Text", "").getText());
var OK = Sys.Process("Maconomy").SWTObject("Shell", "AP Transactions - Invoice Allocation").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
OK.Click();
Delay(2000);
}


  
}


function SOXexcel(CreateClient,start){ 
var Arrayss = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
     Arrayss[id]=temp;
     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrayss;
}

 function excel(CreateClient,start){ 

var Arrayss = [];
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];
Log.Message(DDT.CurrentDriver.ColumnCount);
   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
   xlDriver.Next();

     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=start;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
     if(temp.length!=10){
     Arrayss[id]=temp;
     Log.Message(temp)
     }
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return Arrayss;
}

function ReverseCreditNote(){ 
  gotoMenu();
  invoiceAllocation();
}