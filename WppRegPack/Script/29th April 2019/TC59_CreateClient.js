﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "CreateClient";

var STIME = "";
var GCD1 = [];
var GCD2 = [];
var IND_Specification = []
var SOX_Array = [];
var approvers ="";
var Approve_Level = [];
var LoginArrays = [];
var LoginArr = [];
var LoginEmp = [];
var y = 0;
var OKcount = 0;
var HRData = [];
var UserPasswd = [];
var ClientNo="";
var ifGotIT = true;
//Go To Menu and check for Account Receivable
function GotoMenu(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
    menuBar.DblClick();
    if(ImageRepository.ImageSet.Account_Receivable.Exists()){
ImageRepository.ImageSet.Account_Receivable.Click();
}
else if(ImageRepository.ImageSet.Acc_Receivable_1.Exists()){
ImageRepository.ImageSet.Acc_Receivable_1.Click();
}
else if(ImageRepository.ImageSet.Acc_Receivable_2.Exists()){
ImageRepository.ImageSet.Acc_Receivable_2.Click();  
}


var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|Client Management");
}

}

//var Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 9).SWTObject("Tree", "");
//var Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 10).SWTObject("Tree", "");
//Client_Managt.DblClickItem("|Client Management");
    }
    
    
function Create_Client(){ 
Delay(7000);

var New_Client = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WaitSWTObject("Composite", "",1,60000).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 1);
Sys.HighlightObject(New_Client);

var Add_Visible = true;
while(Add_Visible){
if(New_Client.isEnabled()){
Delay(2000);
Add_Visible = false;  
var New_Client = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
Sys.HighlightObject(New_Client);
New_Client.Click();

Delay(3000);
var confrm = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 26).SWTObject("McPopupPickerWidget", "", 2);
confrm.Keys("Yes");
Delay(3000);

var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
Sys.HighlightObject(next);
next.Click();
}
}
}

//Providing Input to the SOX Compliance for Creating Client
function SOX_Compliance(){ 
SOX_Array = [];
 SOX_Array = SOXexcel(sheetName,1);
 
 var client_identification = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
  if(SOX_Array[0]!=""){
  client_identification.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[0])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
// client_identification.Keys(SOX_Array[0]);
 
 
 Delay(3000);
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[1]!=""){
 checks_did_you_perform.setText(SOX_Array[1]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var new_client_business = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[2]!=""){
 new_client_business.setText(SOX_Array[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var company_owner = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McPopupPickerWidget", "", 2);
   if(SOX_Array[3]!=""){
  company_owner.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[3])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }

 
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[4]!=""){
 checks_did_you_perform.setText(SOX_Array[4]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var foreign_jurisdictions = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McPopupPickerWidget", "", 2);
   if(SOX_Array[5]!=""){
  foreign_jurisdictions.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[5])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[6]!=""){
 checks_did_you_perform.setText(SOX_Array[6]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var sanction_lists = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2);
   if(SOX_Array[7]!=""){
  sanction_lists.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[7])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[8]!=""){
 checks_did_you_perform.setText(SOX_Array[8]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var potential_client_conflicts = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McPopupPickerWidget", "", 2);
   if(SOX_Array[9]!=""){
  potential_client_conflicts.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[9])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
  
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[10]!=""){
 checks_did_you_perform.setText(SOX_Array[10]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var new_client_can_pay = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McPopupPickerWidget", "", 2);
   if(SOX_Array[11]!=""){
  new_client_can_pay.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(SOX_Array[11])
    }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[12]!=""){
 checks_did_you_perform.setText(SOX_Array[12]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }
  Delay(2000);
 var services_provided_new_client = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 23).SWTObject("McTextWidget", "", 2);
   if(SOX_Array[13]!=""){
 services_provided_new_client.setText(SOX_Array[13]);
     }else{ 
    ValidationUtils.verify(false,true,"Needed Data for All Mandatory Fields in SOX Compliance Questions");
  }

  var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
  Sys.HighlightObject(next);
next.Click();
}

//Providing Input to the Global Client Data 1 for Creating Client
function Global_client_Data_1(){ 
GCD1=[];
  GCD1 = SOXexcel(sheetName,3);
var Client_name = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
   if(GCD1[0]!=""){
 Client_name.setText(GCD1[0]+" "+STIME);
     }else{ 
    ValidationUtils.verify(false,true,"Client Name Needed to create Client in Global Client Data 1/2");
  }
//Client_name.Keys(GCD1[0]);
var street1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
   if(GCD1[1]!=""){
 street1.setText(GCD1[1]);
     }else{ 
    ValidationUtils.verify(false,true,"Street1 Needed to create Client in Global Client Data 1/2");
  }
  
//street1.Keys(GCD1[1]);
var street2 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
   if(GCD1[2]!=""){
 street2.setText(GCD1[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Street2 Needed to create Client in Global Client Data 1/2");
  }
  
//street2.Keys(GCD1[2]);

var street3 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
   if(GCD1[3]!=""){
 street3.setText(GCD1[3]);
     }
//street3.Keys(GCD1[3]);
var Area = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
   if(GCD1[4]!=""){
 Area.setText(GCD1[4]);
     }
//Area.Keys(GCD1[4]);
var Postal_code = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
   if(GCD1[5]!=""){
 Postal_code.setText(GCD1[5]);
     }
//Postal_code.Keys(GCD1[5]);
Delay(2000);
var Postal_District = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
   if(GCD1[6]!=""){
 Postal_District.setText(GCD1[6]);
     }
//Postal_District.Keys(GCD1[6]);
var country = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[7]!=""){
  country.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[7])
  }else{ 
    ValidationUtils.verify(false,true,"Country is Needed to create Client in Global Client Data 1/2");
  }
//country.setText(GCD1[7]);
//Delay(5000);
var language = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[8]!=""){
  language.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[8])
  }else{ 
    ValidationUtils.verify(false,true,"Language is Needed to create Client in Global Client Data 1/2");
  }
//language.Keys(GCD1[8]);
//Delay(3000);
var Tax_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
   if(GCD1[9]!=""){
 Tax_No.setText(GCD1[9]);
     }else{ 
    ValidationUtils.verify(false,true,"Tax No Needed to create Client in Global Client Data 1/2");
  }
var Compy_Reg_no = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
   if(GCD1[10]!=""){
 Compy_Reg_no.setText(GCD1[10]);
     }else{ 
    ValidationUtils.verify(false,true,"Company Registration No Needed to create Client in Global Client Data 1/2");
  }
  
var currency = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[11]!=""){
  currency.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[11])
  }else{ 
    ValidationUtils.verify(false,true,"Currency is Needed to create Client in Global Client Data 1/2");
  }
var client_grp = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[12]!=""){
  client_grp.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[12])
  }else{ 
    ValidationUtils.verify(false,true,"Client Group is Needed to Create Client in Global Client Data 1/2");
  }

var control_Acc = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[13]!=""){
  control_Acc.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[13])
  }else{ 
    ValidationUtils.verify(false,true,"Control Account is Needed to Create Client in Global Client Data 1/2");
  }


var party_BFC = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McValuePickerWidget", "", 2);
if(GCD1[14]!=""){
  party_BFC.Click();
  WorkspaceUtils.SearchByValue(party_BFC,"Counter Party BFC",GCD1[14]);
    }else{ 
    ValidationUtils.verify(false,true,"Counter Party BFC is Needed to Create Client in Global Client Data 1/2");
  }

var moda_code = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McValuePickerWidget", "", 2);
//moda_code.Click();
if(GCD1[15]!=""){
  moda_code.Click();
  WorkspaceUtils.SearchByValue(moda_code,"Option",GCD1[15]);
    }else{ 
    ValidationUtils.verify(false,true,"Moda Code is Needed to Create Client in Global Client Data 1/2");
  }

var parent_client = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McValuePickerWidget", "", 2);
if(GCD1[16]!=""){
  parent_client.Click();
  WorkspaceUtils.SearchByValue(parent_client,"Option",GCD1[16]);
    }else{ 
    ValidationUtils.verify(false,true,"Parent Client is Needed to Create Client in Global Client Data 1/2");
  }


 if(GCD1[17]!=""){   
//var invoice_spc_Add = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
 var invoice_spc_Add = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McPopupPickerWidget", "", 2); 
if(invoice_spc_Add.getText()!=GCD1[17]){
 invoice_spc_Add.Click();
 Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible7 = true;
while(Add_Visible7){
if(list.isEnabled()){
Add_Visible7 = false;
    for(var i=list.getItemCount()-1;i>=0;i--){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==GCD1[17]){ 
          list.Keys("[Enter]");

          Delay(5000);
          break;
        }else{ 
          list.Keys("[Up]");
        }
          
      }else{ 
        list.Keys("[Up]");
      }
    }
}
}
}
  Delay(5000);
}

//invoice_spc_Add.Keys("No");
//Delay(4000);

var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
next.Click();
}

//Providing Input to the Global Client Data 2 for Creating Client

function Global_client_Data_2(){
GCD2 = [];

GCD2 = SOXexcel(sheetName,5);
var Company_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
if(GCD2[0]!=""){
  Company_No.Click();
  WorkspaceUtils.SearchByValue(Company_No,"Company",GCD2[0]);
    }else{ 
    ValidationUtils.verify(false,true,"Company Number is Needed to Create Client in Global Client Data 2/2");
  }



var Attn = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
if(GCD2[1]!=""){

    

Attn.Click();

  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.Keys("[Tab]");
    code.setText(GCD2[1]);
    code.Keys("[Down]");
//    
//    code.setText("*india*");
    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(7000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    Log.Message(table.getItemCount());
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==GCD2[1]){ 
        Sys.Desktop.KeyDown(0x28); // Down Arrow
        Delay(1000);
        Sys.Desktop.KeyUp(0x28); 
        Sys.Desktop.KeyDown(0x0D);
        Sys.Desktop.KeyUp(0x0D);
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          Attn.setText("");
        }
      }
      
      }
      Sys.Desktop.KeyUp(0x28);
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
          Attn.setText("");
    }
}else{ 
    ValidationUtils.verify(false,true,"Attn is Needed to Create Client in Global Client Data 2/2");
  }

var Email = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);  
if(GCD2[2]!=""){
 Email.setText(GCD2[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Email Needed to create Client in Global Client Data 2/2");
  }
//Email.Keys(GCD2[2]);
var Phone = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2); 
if(GCD2[3]!=""){
 Phone.setText(GCD2[3]);
     }else{ 
    ValidationUtils.verify(false,true,"Phone Needed to create Client in Global Client Data 2/2");
  }
//Phone.Keys(GCD2[3]);
var Fax = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2); 
if(GCD2[4]!=""){
 Fax.setText(GCD2[4]);
     }
//Fax.Keys(GCD2[4]);
var Acct_Director_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[3]!=""){
Acct_Director_No.Click();

  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(GCD2[5]);
    Delay(3000);
    code.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
    ImageRepository.ImageSet.sale_dropDown.Click();
var code = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);
    code.Keys("Yes");
    Delay(2000);
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    Log.Message(table.getItemCount());
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
    Log.Message("7th Column :"+table.getItem(i).getText_2(7));
   if (table.getItem(i).getText_2(7)!=null){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==GCD2[5]) { 
        Sys.Desktop.KeyDown(0x28); // Down Arrow
        Delay(1000);
        Sys.Desktop.KeyUp(0x28); 
        Sys.Desktop.KeyDown(0x0D);
        Sys.Desktop.KeyUp(0x0D);
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          Acct_Director_No.setText("");
        }
      }
      
      }
      else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
          Acct_Director_No.setText("");
    }
      }
      Sys.Desktop.KeyUp(0x28);
    }

    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
          Acct_Director_No.setText("");
    }
     }else{ 
    ValidationUtils.verify(false,true,"Accountt Director No Needed to create Client in Global Client Data 2/2");
  }


var Account_Manager_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[6]!=""){
Account_Manager_No.Click();

  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(GCD2[6]);
    Delay(3000);
    code.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
    ImageRepository.ImageSet.sale_dropDown.Click();
var code = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);
    code.Keys("Yes");
    Delay(2000);
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    Log.Message(table.getItemCount());
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
    Log.Message("7th Column :"+table.getItem(i).getText_2(7));
   if (table.getItem(i).getText_2(7)!=null){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==GCD2[6]) { 
        Sys.Desktop.KeyDown(0x28); // Down Arrow
        Delay(1000);
        Sys.Desktop.KeyUp(0x28); 
        Sys.Desktop.KeyDown(0x0D);
        Sys.Desktop.KeyUp(0x0D);
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          Account_Manager_No.setText("");
        }
      }
      
      }
      }
      Sys.Desktop.KeyUp(0x28);
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
          Account_Manager_No.setText("");
    }
         }else{ 
    ValidationUtils.verify(false,true,"Accountt Director No Needed to create Client in Global Client Data 2/2");
  }


var Budget_Holder = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McValuePickerWidget", "", 2); 
//Budget_Holder.Click();
if(GCD2[7]!=""){
  Budget_Holder.Click();
  WorkspaceUtils.SearchByValue(Budget_Holder,"Employee",GCD2[7]);
  } 
   
var Main_Biller = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[8]!=""){
  Main_Biller.Click();
  WorkspaceUtils.SearchByValue(Main_Biller,"Employee",GCD2[8]);
  }
var Client_Finance = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[9]!=""){
  Client_Finance.Click();
  WorkspaceUtils.SearchByValue(Client_Finance,"Employee",GCD2[9]);
  }

var Client_Payment_Mode = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[10]!=""){
  Client_Payment_Mode.Click();
  WorkspaceUtils.SearchByValue(Client_Payment_Mode,"Client Payment Mode",GCD2[10]);
         }else{ 
    ValidationUtils.verify(false,true,"Client Payment Mode Needed to create Client in Global Client Data 2/2");
  }


var Payment_Terms = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2); 
if(GCD2[11]!=""){
  Payment_Terms.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(GCD2[11]); 
     }else{ 
    ValidationUtils.verify(false,true,"Payment Terms Needed to create Client in Global Client Data 2/2");
  } 


var Company_Tax_Code = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2); 
if(GCD2[12]!=""){
  Company_Tax_Code.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(GCD2[12]); 
     }else{ 
    ValidationUtils.verify(false,true,"Company Tax Code Needed to create Client in Global Client Data 2/2");
  }


var Level_1_Tax_Derivation = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[13]!=""){
  Level_1_Tax_Derivation.Click();
  WorkspaceUtils.SearchByValue(Level_1_Tax_Derivation,"Local Specification 6",GCD2[13]);
         }else{ 
    ValidationUtils.verify(false,true,"Level 1 Tax Derivation Needed to create Client in Global Client Data 2/2");
  }


var Client_Specific_Logo = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McPopupPickerWidget", "", 2); 
if(GCD2[14]!=""){
  Client_Specific_Logo.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(GCD2[14]); 
     }
//Client_Specific_Logo.Keys(GCD2[11]);



var Job_Surcharge_Rule = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[15]!=""){
  Job_Surcharge_Rule.Click();
  WorkspaceUtils.SearchByValue(Job_Surcharge_Rule,"Job Surcharge Rule",GCD2[15]);
         }

var Job_Price_List_Sales = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[16]!=""){
  Job_Price_List_Sales.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Sales,"Job Price List",GCD2[16]);
         }else{ 
    ValidationUtils.verify(false,true,"Job Price List Sales Needed to create Client in Global Client Data 2/2");
  }


var Job_Price_List_Intercomp = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[17]!=""){
  Job_Price_List_Intercomp.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Intercomp,"Job Price List",GCD2[17]);
         }else{ 
    ValidationUtils.verify(false,true,"Job Price List Intercomp Needed to create Client in Global Client Data 2/2");
  }

var Job_Price_List_Cost = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[18]!=""){
  Job_Price_List_Cost.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Cost,"Job Price List",GCD2[18]);
         }else{ 
    ValidationUtils.verify(false,true,"Job_Price_List_Cost Needed to create Client in Global Client Data 2/2");
  }
  
var Job_Price_List_Standard_Sales = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[19]!=""){
  Job_Price_List_Standard_Sales.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Standard_Sales,"Job Price List",GCD2[19]);
         }else{ 
    ValidationUtils.verify(false,true,"Job_Price_List_Standard_Sales Needed to create Client in Global Client Data 2/2");
  }

var Default_Brand = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 21).SWTObject("McTextWidget", "", 2);
if(GCD2[20]!=""){
  Default_Brand.setText(GCD2[20]+" "+STIME);
         }
  else{ 
        ValidationUtils.verify(false,true,"Default_Brand Needed to create Client in Global Client Data 2/2");
  }
//Default_Brand.Keys(GCD2[16]);
var Default_Product = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 22).SWTObject("McTextWidget", "", 2);
if(GCD2[21]!=""){
  Default_Product.setText(GCD2[21]+" "+STIME);
         }
  else{ 
        ValidationUtils.verify(false,true,"Default_Product Needed to create Client in Global Client Data 2/2");
  }
//Default_Product.Keys(GCD2[17]);
var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
Sys.HighlightObject(next); 
next.Click();

Delay(3000);

//var CICancel = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
//CICancel.Click();
var CIOk = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
CIOk.Click();
Delay(2000);
if(ImageRepository.ImageSet.OK_Button.Exists()){ 
  ImageRepository.ImageSet.OK_Button.Click();
}

Delay(4000);
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Inactive Customers");
inactive.Click();
Delay(5000);
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
ClientName.Keys(GCD1[0]+" "+STIME);
//ClientName.Keys("Automation Client 19December2018 17:32:37");
Delay(5000);

var cltTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var flag=false;
  for(var v=0;v<cltTable.getItemCount();v++){ 
    if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==(GCD1[0]+" "+STIME)){ 
//      if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==SOX_Array[0]){ 
      flag=true;
      ClientNo = cltTable.getItem(v).getText(0).OleValue.toString().trim();
      ValidationUtils.verify(true,true,"Global Client is Available in Maconomy");
//      ReportUtils.logStep("INFO","Global Client is Available in Maconomy");
//      cltTable.Keys("[Enter]");
      break;
    }else{ 
    cltTable.Keys("[Down]");
  }
      
  }
//ClientNo = cltTable.getItem(0).getText(0).OleValue.toString().trim();
if(flag){
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();

if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
  Delay(3000);
}
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var indiaSpecBAR = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
indiaSpecBAR.Click();
Delay(3000);
if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
Delay(3000);
}

IND_Specification = [] ;
IND_Specification = SOXexcel(sheetName,7);
var indiaSpec = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
indiaSpec.Click();
var state_code = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);

if(IND_Specification[0]!=""){
  state_code.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(IND_Specification[0]); 
     }else{ 
    ValidationUtils.verify(false,true,"State Code Needed to create Client in Global Client Data 2/2");
  }

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//var GST_Decor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2);
var GST_Decor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2)
if(IND_Specification[1]!=""){
  GST_Decor.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(IND_Specification[1]); 
     }else{ 
    ValidationUtils.verify(false,true,"GST Debtor Type Needed to create Client in Global Client Data 2/2");
  }

var IndiaPAN = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2)
if(IND_Specification[2]!=""){
IndiaPAN.setText(IND_Specification[2]);
}
Delay(3000);
var IndiaTAN = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
if(IND_Specification[3]!=""){
IndiaTAN.setText(IND_Specification[3]);
}
Delay(3000);
var sav = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
Sys.HighlightObject(sav);
sav.Click();
var barClose = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
Sys.HighlightObject(barClose);
barClose.Click();
Delay(3000);



var document = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
document.Click();
Delay(3000);

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var atthDoc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
atthDoc.Click();
Delay(4000);
var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
dicratory.Keys("C:\\Users\\674087\\Desktop\\New folder\\test1.xlsx");
var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
Sys.HighlightObject(opendoc);
opendoc.Click();


var home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
home.Click();
Delay(3000);
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
info.Click();
Delay(5000);
var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7)
Sys.HighlightObject(submit);
submit.Click();
Delay(2000);
if(ImageRepository.ImageSet.OK_Button.Exists()){
ImageRepository.ImageSet.OK_Button.Click();
}
Delay(2000);
if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();

}
Delay(3000);
var AllApproved = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 9);
AllApproved.Click();
Delay(4000);
y =0 ;
//var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
for(var i=0;i<ApproverTable.getItemCount();i++){ 
  
    approvers="";
       approvers = ClientNo+"*"+GCD1[0]+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +i+ ": " +approvers);
//       Approve_Level[y] = Company_ID+"*"+Job_Name+"*"+approvers;
Approve_Level[y] = approvers;
       y++;
}
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "").Click();
Delay(3000);
if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
}
//goToHR();
//Credentiallogin();
//if(GCD2[0]!="")
//UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,GCD2[0]);
//else
//UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,"");
//
//RestMaconomy(UserPasswd)
}
}

function vv(){
//goToHR();
//Credentiallogin();
//var Approve_Level = [];
////Approve_Level[0]="122232*Automation Client 19December2018 19:53:38*1707 - Senior Finance (170710118)*1707 - Management (170710127)";
////Approve_Level[1]="122232*Automation Client 19December2018 19:53:38*SACHINDRA P KARKERA (170710011)*";
////Approve_Level[2]="122232*Automation Client 19December2018 19:53:38*Central Team - Client Management*Central Team - Vendor Management";
//
//UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,"1707");
//RestMaconomy(UserPasswd)
}





function SOXexcel(CreateClient,start){ 
//function SOXexcel(){ 
 var Arrays = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
//var xlDriver = DDT.ExcelDriver(Project.Path+excelName, "CreateClient", true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
//      for(var idx=1;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
//      }
     Arrays[id]=temp;
     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}

//While Creating Maconomy should not Maximize
//After Client Created India Specification Bar should be "minimize";
function CreateClient(){ 
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Create Client test started::"+STIME);
  GotoMenu();
  Create_Client();
  SOX_Compliance();
  Global_client_Data_1();
  Global_client_Data_2();
  WorkspaceUtils.closeAllWorkspaces();
}




function Credentiallogin() {
  var xlDriver = DDT.ExcelDriver(Project.Path+excelName, "userRoles", false);
var id =0;
var colsList = [];

 for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
   colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
 }
   while (!DDT.CurrentDriver.EOF()) {
   var temp ="";
    for(var idx=0;idx<colsList.length;idx++){  
     if(xlDriver.Value(colsList[idx])!=null){
    temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
    }
    else{ 
      temp = temp+"*";
    }
    }
//      Log.Message(temp)
   LoginEmp[id]=temp;
   id++;     
   xlDriver.Next();
   }
   DDT.CloseDriver(xlDriver.Name);
}



function RestMaconomy(UserPasswd){ 
//var UserPasswd = [];
//UserPasswd[0] = "122219*Regular Hindustan*1710 - Finance*CORE@WPP123";
//UserPasswd[1] = "122219*Regular Hindustan*somsubhra.banerjee@jwt.com*CORE@WPP123";
//UserPasswd[2] = "122219*Regular Hindustan*SSC IN -  CT Clients*CORE@WPP123";
Log.Message(UserPasswd.length);
for(var i=0;i<UserPasswd.length;i++){

var temp = UserPasswd[i];
var temp_user = temp.split("*");
var uname = temp_user[2]; 
Log.Message(uname)
var pwd = temp_user[3];
Log.Message(pwd)
Rests(uname,pwd);
    
GotoMenu();

Delay(7000);
var stat = true;
while(stat){
 var ref_Image = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 1); 
  if(ref_Image.isEnabled()){ 
    stat = false;
  }else{ 
    Delay(2000);
  }
}
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Inactive Customers");
inactive.Click();
Delay(5000);
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
ClientNo.Keys(temp_user[0]);
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//ClientName.Keys(GCD1[0]+" "+STIME);
ClientName.Keys(temp_user[1]);
Delay(5000);
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();
Delay(3000);
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//if(ImageRepository.ImageSet.Apprve.Exists()){ 
//  ImageRepository.ImageSet.Apprve.Click();
//}
//else{ 
//Log.Warning("Client is not Approved by :"+uname);
//}
Delay(5000);
var approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(approve)
if(approve.isEnabled()){ 
 approve.Click(); 
 if(OKcount==2)
 for(var j=0;j<12;j++){ 
    if(ImageRepository.ImageSet.Ok.Exists()){ 
     ImageRepository.ImageSet.Ok.Click();
     Delay(1000);
   }
   else if(ImageRepository.ImageSet.OK_Button.Exists()){ 
     ImageRepository.ImageSet.OK_Button.Click();
     Delay(1000);
   }
 }
 OKcount++;
 ValidationUtils.verify(true,true,"Global Client is Approved by "+uname);
}
else{ 
Log.Warning("Client is not Approved by :"+uname);
}

}


}

//function Login_Match(){ 
//Delay(3000);
//login();
//logins();
//goToHR();
//Credentiallogin();
//var z =0;
//for(var i=0;i<Approve_Level.length;i++){ 
//Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,GCD2[0]);
//// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
//  if(Approve_Level[i].indexOf("SSC - Biller")==-1){
//  Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
//  }
//
//var tempLevel = Approve_Level[i].split("*");
//ifGotIT = true;
//for(var j=2;j<tempLevel.length;j++){ 
//
//if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
//var temp = tempLevel[j].replace(" (","*");
//temp = temp.replace(")","");
////Log.Message("temp :"+temp)
//var tempSplit = temp.split("*");
//
//  for(var k=0;k<LoginEmp.length;k++){
//    var A_temp = LoginEmp[k].split("*");
////    Log.Message("tempSplit[0] :"+tempSplit[0]);
////    Log.Message("A_temp[0] :"+A_temp[0]);
////    Log.Message("tempSplit[1] :"+tempSplit[1]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
//   if((tempSplit[0]==A_temp[0]) || (tempSplit[1]==A_temp[1])){ 
//      UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
//     Log.Message(UserPasswd[z]);
//     z++;
//     ifGotIT = false;
//     break;     
//   }
//    
//  }
//  if(!ifGotIT){ 
//    break;
//  }
//  }
//  
//  if((tempLevel[j].indexOf("SSC -")!=-1) || (tempLevel[j].indexOf("Central Team -")!=-1)){ 
//  for(var k=0;k<LoginArrays.length;k++){
//    var A_temp = LoginArrays[k].split("*");
////    Log.Message("tempSplit[j] :"+tempLevel[j]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
//    if(A_temp[1].indexOf("Central Team - Client")!=-1){ 
//      A_temp[1] = "Central Team - Client Management";
//    }
//    if(A_temp[1].indexOf("Central Team - Vendor")!=-1){ 
//      A_temp[1] = "Central Team - Vendor Management";
//    }
//    
//   if(tempLevel[j]==A_temp[1]){ 
//     UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     
//     Log.Message(UserPasswd[z]);
//     z++;
//     ifGotIT = false;
//     break;     
//   }
//    
//  }
//  if(!ifGotIT){ 
//    break;
//  }
//  }
//  
//if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
//var temp = tempLevel[j].replace(" (","*");
//temp = temp.replace(")","");
////Log.Message("temp :"+temp)
//var tempSplit = temp.split("*");
//
//  for(var k=0;k<HRData.length;k++){
//    var A_temp = HRData[k].split("*");
////    Log.Message("tempSplit[0] :"+tempSplit[0]);
////    Log.Message("A_temp[0] :"+A_temp[0]);
////    Log.Message("tempSplit[1] :"+tempSplit[1]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
//   if(tempSplit[1]==A_temp[1]){ 
//     UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
//     Log.Message(UserPasswd[z]);
//     z++;
//     ifGotIT = false;
//     break;     
//   }
//    
//  }
//  if(!ifGotIT){ 
//    break;
//  }
//  }
//  if((tempLevel[j].indexOf(" (")==-1) && (tempLevel[j].indexOf(")")==-1) && 
//  (tempLevel[j].indexOf("SSC -")==-1) && (tempLevel[j].indexOf("Central Team -")==-1)){ 
//    
//  for(var k=0;k<LoginArr.length;k++){
//  var A_temp = LoginArr[k].split("*");
//    if(A_temp[0]==tempLevel[j]){  // Better  to use level[j].indexOf(LoginArrays[k])
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[1]; 
//    Log.Message(UserPasswd[z]);
//     z++;
//     ifGotIT = false;
//     break;     
//   }
//   }
//  if(!ifGotIT){ 
//    break;
//  }
//  }
//  
//  }
//  if(ifGotIT){ 
//    Log.Warning("UserName and Password is Not Matched for Approver and Substitute :"+Approve_Level[i]);
//  }
//  
//}
//
//
//}

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
    }
    
    
function goToHR(){ 
Delay(3000);
  closeAllWorkspaces();
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();

if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();  
}

var HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
HRitem.DblClickItem("|Users");
Delay(5000);
//var ActiveUser = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Users");
//ActiveUser.Click();
var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click();
Delay(5000);
var HRTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var z=0;
for(var i=0;i<HRTable.getItemCount();i++){ 
if(HRTable.getItem(i).getText(2)!=""){
HRData[z] = HRTable.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+HRTable.getItem(i).getText_2(2).OleValue.toString().trim()
//Log.Message(HRData[z]);
z++;

}
}

}


function login() {
    var xlDriver = DDT.ExcelDriver(workBook, sscCredential, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
//      Log.Message(temp)
     LoginArrays[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}


//function Credentiallogin() {
//    var xlDriver = DDT.ExcelDriver(workBook, Credential, true);
//var id =0;
//var colsList = [];
//
//   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
//     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
//   }
//     while (!DDT.CurrentDriver.EOF()) {
//     var temp ="";
//      for(var idx=0;idx<colsList.length;idx++){  
//       if(xlDriver.Value(colsList[idx])!=null){
//      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
//      }
//      else{ 
//        temp = temp+"*";
//      }
//      }
////      Log.Message(temp)
//     LoginEmp[id]=temp;
//     id++;     
//     xlDriver.Next();
//     }
//     DDT.CloseDriver(xlDriver.Name);
//}


function Rests(uname,pwd){ 
Delay(5000);
      Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x52); //R 
     Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
     Sys.Desktop.KeyUp(0x52); //R
Delay(65000);
     var usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
    var pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    var btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
    usernameAddr.SetFocus();
    usernameAddr.setText(uname);
    pwdAddr.setText(pwd);
    btnLogin.click();
    Delay(10000);   
}



function logins() {
    var xlDriver = DDT.ExcelDriver(workBook, loginpassword, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
//      Log.Message(temp)
     LoginArr[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}
