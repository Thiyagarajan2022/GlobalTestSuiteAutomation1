﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "CreateVendor";
var venDetails = [];
var venDetails2 = [];
var SOX_Array = [];

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
Client_Managt.DblClickItem("|Vendor Management");
}

}
}

function gotoVendorCreation(){ 
Delay(7000);
  venDetails = SOXexcel(sheetName,1)
  var companyNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
  if(venDetails[0]!=""){
  companyNo.Click();
  WorkspaceUtils.SearchByValue(companyNo,"Company",venDetails[0]);
    }else{ 
    ValidationUtils.verify(false,true,"Company Number is Need to create");
  }
  var vendorName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
  vendorName.Click();
 if(venDetails[1]!=""){
 vendorName.Keys("^a[BS]");
 vendorName.setText(venDetails[1]+" "+STIME);
 }else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Need to create");
  }
  var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  save.Click();
  Delay(4000);
  var global = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  global.Click();
  Delay(3000);
  var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Inactive Vendors");
  inactive.Click();
  var NewVendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  NewVendor.Click();
}




function vendorDetails(){ 
Delay(4000);
  venDetails = [];
  venDetails = SOXexcel(sheetName,1)
  var vendorName = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
     if(venDetails[1]!=""){
 vendorName.setText(venDetails[1]+" "+STIME);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var street1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  if(venDetails[2]!=""){
 street1.setText(venDetails[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var street2 = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  if(venDetails[3]!=""){
 street2.setText(venDetails[3]);
     }
     
  var street3 = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
  if(venDetails[4]!=""){
 street3.setText(venDetails[4]);
     }
     
  var area = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
  if(venDetails[5]!=""){
 area.setText(venDetails[5]);
     }
     
  var postalCode = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
  if(venDetails[6]!=""){
 postalCode.setText(venDetails[6]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var city = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(venDetails[7]!=""){
 city.setText(venDetails[7]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var country = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails[8]!=""){
  country.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails[8])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var language = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails[9]!=""){
  language.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails[9])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var taxNo = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2);
  if(venDetails[10]!=""){
 taxNo.setText(venDetails[10]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var CompyRegNo = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
  if(venDetails[11]!=""){
 CompyRegNo.setText(venDetails[11]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var Currency = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails[12]!=""){
  Currency.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails[12])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var vendorGroup = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails[13]!=""){
  vendorGroup.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails[13])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var controlAccount = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails[14]!=""){
  controlAccount.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails[14])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var trade = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails[15]!=""){
  trade.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails[15])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var partyBFC = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McValuePickerWidget", "", 2);
  if(venDetails[16]!=""){
  partyBFC.Click();
  WorkspaceUtils.SearchByValue(partyBFC,"Counter Party BFC",venDetails[16]);
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var parentVendor = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McValuePickerWidget", "", 2);
  if(venDetails[17]!=""){
  parentVendor.Click();
  WorkspaceUtils.SearchByValue(parentVendor,"Option",venDetails[17]);
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var BankName = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McTextWidget", "", 2);
  if(venDetails[18]!=""){
 BankName.setText(venDetails[18]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var IBAN = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McTextWidget", "", 2);
  if(venDetails[19]!=""){
 IBAN.setText(venDetails[19]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var AccNumber = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McTextWidget", "", 2);
  if(venDetails[20]!=""){
 AccNumber.setText(venDetails[20]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var sortCode = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McTextWidget", "", 2);
  if(venDetails[21]!=""){
 sortCode.setText(venDetails[21]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var swift = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 21).SWTObject("McTextWidget", "", 2);
  if(venDetails[22]!=""){
 swift.setText(venDetails[22]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Global Vendor Master Data");
  }
  
  var intermidiate = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 22).SWTObject("McTextWidget", "", 2);
  if(venDetails[23]!=""){
 intermidiate.setText(venDetails[23]);
     }
     
  Delay(3000);
  var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
  next.Click();
//  var create = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
  

}

function vendorDetails2(){ 
Delay(4000);
  venDetails2 = [];
  venDetails2 = SOXexcel(sheetName,3);
  var companyNo = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
  if(venDetails2[0]!=""){
  companyNo.Click();
  WorkspaceUtils.SearchByValue(companyNo,"Company",venDetails2[0]);
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var Attn = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  if(venDetails2[1]!=""){
 Attn.setText(venDetails2[1]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var Email = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  if(venDetails2[2]!=""){
 Email.setText(venDetails2[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var Remitted_Email = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
  if(venDetails2[3]!=""){
 Remitted_Email.setText(venDetails2[3]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var phone = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
  if(venDetails2[4]!=""){
 phone.setText(venDetails2[4]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var fax = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
  if(venDetails2[5]!=""){
 fax.setText(venDetails2[5]);
     }
     
  var paymentTerm = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails2[6]!=""){
  paymentTerm.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails2[6])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var companyTaxCode = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
   if(venDetails2[7]!=""){
  companyTaxCode.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(venDetails2[7])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var Leve1TaxDes = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McValuePickerWidget", "", 2);
  if(venDetails2[8]!=""){
  Leve1TaxDes.Click();
  WorkspaceUtils.SearchByValue(Leve1TaxDes,"Local Specification 6",venDetails2[8]);
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  
  var paymentMode = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McValuePickerWidget", "", 2);
  if(venDetails2[9]!=""){
  paymentMode.Click();
  WorkspaceUtils.SearchByValue(paymentMode,"Payment Mode",venDetails2[9]);
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Local Vendor Master Data");
  }
  Delay(3000);
  var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
  next.Click();
  Delay(3000);
  var policyConfirm = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 35).SWTObject("McPopupPickerWidget", "", 2);
  policyConfirm.Keys("Yes");
  Delay(3000);
  var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
  next.Click();
  Delay(3000);
   
}

function DiligenceCheckout(){ 
  
}


function vendorDetails3(){ 
 SOX_Array = [];
 SOX_Array = SOXexcel(sheetName,5);
 Delay(3000);
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[0]!=""){
 checks_did_you_perform.setText(SOX_Array[0]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
  var WPPagencyPreferredSupplier = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2);
  if(SOX_Array[1]!=""){
  WPPagencyPreferredSupplier.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[1])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
   var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[2]!=""){
 checks_did_you_perform.setText(SOX_Array[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
    var WPPagencyPreferredSupplier = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McPopupPickerWidget", "", 2);
  if(SOX_Array[3]!=""){
  WPPagencyPreferredSupplier.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[3])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
  
  
  
    var satisfiedSupplier = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McPopupPickerWidget", "", 2);
  if(SOX_Array[4]!=""){
  satisfiedSupplier.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[4])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
  
   var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[5]!=""){
 checks_did_you_perform.setText(SOX_Array[5]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
  
    var awarePersonalRelationship = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2);
  if(SOX_Array[6]!=""){
  awarePersonalRelationship.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[6])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
   var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[7]!=""){
 checks_did_you_perform.setText(SOX_Array[7]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
   var valueFirstOrder = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[8]!=""){
 valueFirstOrder.setText(SOX_Array[8]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
   var AnnualSpent = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[9]!=""){
 AnnualSpent.setText(SOX_Array[9]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in Due Diligence Checkout");
  }
  
  Delay(3000);
var create = Sys.Process("Maconomy").SWTObject("Shell", "Create Vendor").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
create.Click();

Delay(4000);
var ok = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Management - Vendor Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
ok.Click();
Delay(4000);
var ok = Sys.Process("Maconomy").SWTObject("Shell", "Vendor Management - Vendor Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
ok.Click();
Delay(3000);
var global = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
global.Click();
Delay(3000);
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Inactive Vendors");
inactive.Click();
Delay(3000);
var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "");
firstcell.Click();
Delay(4000);
firstcell.Keys("[Down]");
Delay(3000);
table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
table.Keys("[Tab]");  
table.Keys("[Enter]");  
// Can't able to click hyperLink once try
}

function SOXexcel(CreateClient,start){ 
var Arrays = []; 
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
     Arrays[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}


function CreateVendor(){ 
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create Client test started::"+STIME);
  gotoMenu();
  gotoVendorCreation();
  vendorDetails();
  vendorDetails2();
  vendorDetails3();
}

function v(){ 
table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)

Log.Message( table.getItem(0).getItem_2(1).FullName);
}