﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "CreateProduct";
var STIME = "";
var GCD1=[];
var GCD2 = [];
var IND_Specification = []
var SOX_Array = [];
var approvers ="";
var Approve_Level = [];
var approvers ="";
var HRData = [];
var LoginEmp = [];
var y = 0;
var UserPasswd = [];
var ClientNo="";
var ifGotIT = true;
var brName = "";
var ClientNumber = "";
var level = 0;

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



//    if(ImageRepository.ImageSet.Clnt_Mangmt.Exists()){
//ImageRepository.ImageSet.Clnt_Mangmt.DblClick();
//}
//else if(ImageRepository.ImageSet.Clnt_Mangmt_2.Exists()){
//ImageRepository.ImageSet.Clnt_Mangmt_2.DblClick();
//}
//else if(ImageRepository.ImageSet.Clnt_Mangmt_3.Exists()){
//ImageRepository.ImageSet.Clnt_Mangmt_2.DblClick();
//}


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


    function searchcompany(){ 

SOX_Array = [];
SOX_Array = SOXexcel(sheetName,1);
Delay(4000);
var stat = true;
while(stat){
 var ref_Image = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 1); 
  if(ref_Image.isEnabled()){ 
    stat = false;
  }else{ 
    Delay(2000);
  }
}

var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Customers");
inactive.Click();
Delay(5000);
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
ClientNo.setText(SOX_Array[1]);
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
ClientName.setText(SOX_Array[0]);
//ClientName.Keys("Automation Client 19December2018 19:42:22");
Log.Message(SOX_Array[0])
ClientName.Keys("[Tab]");
var AppList = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 3);
AppList.Click();
Delay(4000);
ImageRepository.ImageSet.sale_dropDown.Click();
Delay(1000);
Sys.Process("Maconomy").Refresh();
var code = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);
    code.Keys("Yes");
    Delay(5000);
    Sys.Desktop.KeyDown(0x0D);
    Sys.Desktop.KeyUp(0x0D);

Delay(5000);
var cltTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var flag = false;
  for(var v=0;v<cltTable.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(temp_user[1]+" "+STIME)){ 
      if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==SOX_Array[0]){ 
      flag=true;
      ClientNo = cltTable.getItem(v).getText(2).OleValue.toString().trim();
//      ReportUtils.logStep("INFO","Global Client is Available to create Company Client");
      break;
    }else{ 
    cltTable.Keys("[Down]");
  }
      
  }
if(flag){
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();
Delay(5000);
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var subCustomer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
subCustomer.Click();
Delay(5000);
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var createBand = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
createBand.Click();
Delay(4000);
GCD1=[];
GCD1 = SOXexcel(sheetName,3);
var brandName = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
if(GCD1[0]!=""){
brandName.setText(GCD1[0]+" "+STIME);
brName = GCD1[0]+" "+STIME;
}else{ 
ValidationUtils.verify(false,true,"Brand Name Needed to create Product in Global Client");
}
var street1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
if(GCD1[1]!=""){
street1.setText(GCD1[1]);
}else{ 
ValidationUtils.verify(false,true,"Street1 Name Needed to create Product in Global Client");
}
var street2 = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
if(GCD1[2]!=""){
street2.setText(GCD1[2]);
}
var street3 = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2)
if(GCD1[3]!=""){
street3.setText(GCD1[3]);
}
var street4 = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2)
if(GCD1[4]!=""){
street4.setText(GCD1[4]);
}
var postalCode = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
if(GCD1[5]!=""){
postalCode.setText(GCD1[5]);
}
var postalDistrict = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(GCD1[6]!=""){
postalDistrict.setText(GCD1[6]);
}
var Country = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[7]!=""){
  Country.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[7])
  }else{ 
    ValidationUtils.verify(false,true,"Country is Needed to create Product in Global Client");
  }
var language = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2);
  if(GCD1[8]!=""){
  language.Click();

  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList(GCD1[8])
  }else{ 
    ValidationUtils.verify(false,true,"Language is Needed to create Product in Global Client");
  }
var BFC = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McValuePickerWidget", "", 2);
if(GCD1[9]!=""){
  BFC.Click();
  WorkspaceUtils.SearchByValue(BFC,"Counter Party BFC",GCD1[9]);
    }else{ 
    ValidationUtils.verify(false,true,"BFC Number is Needed to Create Product in Global Client Data 2/2");
  }
var modaCode = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McValuePickerWidget", "", 2);
if(GCD1[10]!=""){
  modaCode.Click();
  WorkspaceUtils.SearchByValue(modaCode,"Option",GCD1[10]);
    }else{ 
    ValidationUtils.verify(false,true,"Moda Code is Needed to Create Product in Global Client Data 2/2");
  }
Delay(4000);
var nxt = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
Sys.HighlightObject(nxt);
nxt.Click();
}
}

function Global_client_Data_2(){
GCD2 = [];

GCD2 = SOXexcel(sheetName,5);
var Company_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
if(GCD2[0]!=""){
  Company_No.Click();
  WorkspaceUtils.SearchByValue(Company_No,"Company",GCD2[0]);
    }else{ 
    ValidationUtils.verify(false,true,"Company Number is Needed to Create Product in Global Client Data 2/2");
  }



var Attn = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2)
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
//    code.Keys("[Down]");
   
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

var Email = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2)  
if(GCD2[2]!=""){
 Email.setText(GCD2[2]);
     }else{ 
    ValidationUtils.verify(false,true,"Email Needed to create Product in Global Client Data 2/2");
  }
//Email.Keys(GCD2[2]);
var Phone = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2) 
if(GCD2[3]!=""){
 Phone.setText(GCD2[3]);
     }else{ 
    ValidationUtils.verify(false,true,"Phone Needed to create Product in Global Client Data 2/2");
  }
//Phone.Keys(GCD2[3]);
var Fax = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2) 
if(GCD2[4]!=""){
 Fax.setText(GCD2[4]);
     }
//Fax.Keys(GCD2[4]);
var Acct_Director_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McValuePickerWidget", "", 2) 
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
    ValidationUtils.verify(false,true,"Accountt Director No Needed to create Product in Global Client Data 2/2");
  }


var Account_Manager_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McValuePickerWidget", "", 2) 
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
    ValidationUtils.verify(false,true,"Accountt Director No Needed to create Product in Global Client Data 2/2");
  }


var Budget_Holder = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McValuePickerWidget", "", 2) 
//Budget_Holder.Click();
if(GCD2[7]!=""){
  Budget_Holder.Click();
  WorkspaceUtils.SearchByValue(Budget_Holder,"Employee",GCD2[7]);
  } 
   
var Main_Biller = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[8]!=""){
  Main_Biller.Click();
  WorkspaceUtils.SearchByValue(Main_Biller,"Employee",GCD2[8]);
  }
var Client_Finance = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[9]!=""){
  Client_Finance.Click();
  WorkspaceUtils.SearchByValue(Client_Finance,"Employee",GCD2[9]);
  }

var Client_Payment_Mode = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[10]!=""){
  Client_Payment_Mode.Click();
  WorkspaceUtils.SearchByValue(Client_Payment_Mode,"Client Payment Mode",GCD2[10]);
         }else{ 
    ValidationUtils.verify(false,true,"Client Payment Mode Needed to create Product in Global Client Data 2/2");
  }


var Payment_Terms = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2) 
if(GCD2[11]!=""){
  Payment_Terms.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(GCD2[11]); 
     }else{ 
    ValidationUtils.verify(false,true,"Payment Terms Needed to create Product in Global Client Data 2/2");
  } 


var Company_Tax_Code = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2) 
if(GCD2[12]!=""){
  Company_Tax_Code.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(GCD2[12]); 
     }else{ 
    ValidationUtils.verify(false,true,"Company Tax Code Needed to create Product in Global Client Data 2/2");
  }


var Level_1_Tax_Derivation = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[13]!=""){
  Level_1_Tax_Derivation.Click();
  WorkspaceUtils.SearchByValue(Level_1_Tax_Derivation,"Local Specification 6",GCD2[13]);
         }else{ 
    ValidationUtils.verify(false,true,"Level 1 Tax Derivation Needed to create Product in Global Client Data 2/2");
  }


var Client_Specific_Logo = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McPopupPickerWidget", "", 2) 
if(GCD2[14]!=""){
  Client_Specific_Logo.Click();  
  Delay(5000);
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(GCD2[14]); 
     }
//Client_Specific_Logo.Keys(GCD2[11]);



var Job_Surcharge_Rule = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McValuePickerWidget", "", 2); 
if(GCD2[15]!=""){
  Job_Surcharge_Rule.Click();
  WorkspaceUtils.SearchByValue(Job_Surcharge_Rule,"Job Surcharge Rule",GCD2[15]);
         }

var Job_Price_List_Sales = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[16]!=""){
  Job_Price_List_Sales.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Sales,"Job Price List",GCD2[16]);
         }else{ 
    ValidationUtils.verify(false,true,"Job Price List Sales Needed to create Product in Global Client Data 2/2");
  }


var Job_Price_List_Intercomp = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[17]!=""){
  Job_Price_List_Intercomp.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Intercomp,"Job Price List",GCD2[17]);
         }else{ 
    ValidationUtils.verify(false,true,"Job Price List Intercomp Needed to create Product in Global Client Data 2/2");
  }

var Job_Price_List_Cost = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[18]!=""){
  Job_Price_List_Cost.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Cost,"Job Price List",GCD2[18]);
         }else{ 
    ValidationUtils.verify(false,true,"Job_Price_List_Cost Needed to create Product in Global Client Data 2/2");
  }
  
var Job_Price_List_Standard_Sales = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McValuePickerWidget", "", 2) 
if(GCD2[19]!=""){
  Job_Price_List_Standard_Sales.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Standard_Sales,"Job Price List",GCD2[19]);
         }else{ 
    ValidationUtils.verify(false,true,"Job_Price_List_Standard_Sales Needed to create Product in Global Client Data 2/2");
  }

//var Default_Brand = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 21).SWTObject("McTextWidget", "", 2);
//if(GCD2[20]!=""){
//  Default_Brand.setText(GCD2[20]);
//         }
//  else{ 
//        ValidationUtils.verify(false,true,"Default_Brand Needed to create Client in Global Client Data 2/2");
//  }
//Default_Brand.Keys(GCD2[16]);
var Default_Product = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 21).SWTObject("McTextWidget", "", 2)
Default_Product.getText()
if((GCD2[20]!="") && (GCD2[20]!=Default_Product.getText())){
  Default_Product.setText(GCD2[20]);
         }
  else{ 
        ValidationUtils.verify(false,true,"Default_Product Needed to create Product in Global Client Data 2/2");
  }

var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Product").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
Sys.HighlightObject(next); 
next.Click();

Delay(3000);

//var CICancel = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
//CICancel.Click();
var CIOk = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
CIOk.Click();
Delay(5000);
if(ImageRepository.ImageSet.Ok.Exists()){ 
  ImageRepository.ImageSet.Ok.Click();
}
Delay(5000);
if(ImageRepository.ImageSet.OK_Button.Exists()){ 
  ImageRepository.ImageSet.OK_Button.Click();
}



Delay(5000);
var listSubCustomer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
listSubCustomer.Click();
Delay(5000);
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstcell.Click();
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(1).OleValue.toString().trim()==brName){ 
      flag=true;
      table.Keys("[Down]");
      ClientNumber = table.getItem(v).getText_2(0).OleValue.toString().trim();
      Log.Message(table.getItem(v).getText_2(0).OleValue.toString().trim())
//      table.Keys("[Enter]");
      break;
    }else{ 
    table.Keys("[Down]");
  }
      
  }
  Delay(2000);
  var closefliter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2)
  closefliter.Click();
  Delay(2000);
  WorkspaceUtils.closeAllWorkspaces();
  GotoMenu();
  Delay(6000);


  
  
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Inactive Customers");
inactive.Click();
Delay(6000);
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
ClientNo.setText(ClientNumber);
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
ClientName.setText(brName);
//ClientName.Keys("Automation Client 19December2018 17:32:37");
Delay(5000);
var cltTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var flag=false;
  for(var v=0;v<cltTable.getItemCount();v++){ 
      if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==brName){ 
      flag=true;
//      table.Keys("[Down]");
//      ClientNumber = table.getItem(v).getText_2(0).OleValue.toString().trim();
//      Log.Message(table.getItem(v).getText_2(0).OleValue.toString().trim())
//      table.Keys("[Enter]");
      break;
    }else{ 
    cltTable.Keys("[Down]");
  }
      
  }



//ClientNo = cltTable.getItem(0).getText(0).OleValue.toString().trim();
if(flag){
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
}
var indiaSpecBAR = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
indiaSpecBAR.Click();
Delay(3000);
ImageRepository.ImageSet.Maximize.Click();
Delay(3000);

IND_Specification = [] ;
IND_Specification = SOXexcel(sheetName,7);
var indiaSpec = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
indiaSpec.Click();
var state_code = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
var checkmark = false;
if((IND_Specification[0]!="") && (IND_Specification[0]!=state_code.getText())){
  state_code.Keys("01");
//ApproverGroup.Keys("Default");
Delay(5000);

state_code.Click();

//Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible3 = true;
 while(Add_Visible3){
if(list.isEnabled()){
Add_Visible3 = false;
for(var i=0;i<4;i++){ 
list.Keys("[Up]");
}
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==IND_Specification[0]){ 
          Delay(1000);

            
          list.Keys("[Enter]");
          Delay(3000);
          break;
        }else{ 
        if(i!=0){
         Log.Message(list.getItem(i).getText_2(0));
        list.Keys("[Down]");
          }
        }
          
      }else{ 
      if(i!=0){
         Log.Message(list.getItem(i).getText_2(0));
        list.Keys("[Down]");
          }
        
      }
    }
}
}
  checkmark = true;
  ReportUtils.logStep("INFO", "StateCode has Changed");
  }

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
//var GST_Decor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2);
var GST_Decor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McPopupPickerWidget", "", 2)
if((IND_Specification[1]!="") && (IND_Specification[1]!=GST_Decor.getText())){
  GST_Decor.Keys("-");
//ApproverGroup.Keys("Default");
Delay(5000);

GST_Decor.Click();

//Sys.Process("Maconomy").Refresh();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible3 = true;
 while(Add_Visible3){
if(list.isEnabled()){
Add_Visible3 = false;
for(var i=0;i<4;i++){ 
list.Keys("[Up]");
}
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==IND_Specification[1]){ 
          Delay(1000);

            
          list.Keys("[Enter]");
          Delay(3000);
          break;
        }else{ 
        if(i!=0){
         Log.Message(list.getItem(i).getText_2(0));
        list.Keys("[Down]");
          }
        }
          
      }else{ 
      if(i!=0){
         Log.Message(list.getItem(i).getText_2(0));
        list.Keys("[Down]");
          }
        
      }
    }
}
}
  checkmark = true;
  ReportUtils.logStep("INFO", "GST Decor Type has Changed");
  }

var IndiaPAN = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2)
IndiaPAN.Click();
if((IND_Specification[2]!="")&& (IND_Specification[2]!=IndiaPAN.getText())){
IndiaPAN.setText(IND_Specification[2]);
checkmark = true;
}
Delay(3000);
var IndiaTAN = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2);
IndiaTAN.Click();
if((IND_Specification[3]!="")&& (IND_Specification[3]!=IndiaTAN.getText())){
IndiaTAN.setText(IND_Specification[3]);
checkmark = true;
}
Delay(3000);
if(checkmark){
var sav = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
Sys.HighlightObject(sav);
sav.Click();
if(ImageRepository.ImageSet.Ok.Exists()){ 
  ImageRepository.ImageSet.Ok.Click();
}else if(ImageRepository.ImageSet.OK_Button.Exists()){ 
  ImageRepository.ImageSet.OK_Button.Click();
}
}
Delay(2000);
var barClose = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
Sys.HighlightObject(barClose);
barClose.Click();
//Delay(3000);
//if(ImageRepository.ImageSet.Forward.Exists()){ 
//  ImageRepository.ImageSet.Forward.Click();
//}
Delay(2000);
var document = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
document.Click();
Delay(3000);

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var atthDoc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
if(atthDoc.getText()=="New"){ 
 atthDoc.Click(); 
}else{
var atthDoc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 10);
if(atthDoc.getText()=="New"){ 
 atthDoc.Click(); 
}
}
Delay(4000);
var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
dicratory.Keys("C:\\Users\\674087\\Desktop\\New folder\\test4.xlsx");
var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
Sys.HighlightObject(opendoc);
opendoc.Click();

Delay(4000);
var home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
home.Click();
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
info.Click();
Delay(5000);
var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7)
Sys.HighlightObject(submit);
submit.Click();
ValidationUtils.verify(true,true,"Product has been Created");

Delay(2000);
if(ImageRepository.ImageSet.OK_Button.Exists()){
ImageRepository.ImageSet.OK_Button.Click();
}
Delay(5000);
//var indiaSpecBAR = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
//indiaSpecBAR.Click();
//Delay(3000);

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
       approvers = ClientNumber+"*"+brName+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +i+ ": " +approvers);
//       Approve_Level[y] = Company_ID+"*"+Job_Name+"*"+approvers;
Approve_Level[y] = approvers;
       y++;
}
Delay(2000);


HRData = goToHR();
LoginEmp = Credentiallogin(Project.Path+excelName, "userRoles");

if(GCD2[0]!="")
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,GCD2[0]);
else
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,"");

RestMaconomy(UserPasswd)
}
}


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
 ValidationUtils.verify(true,true,"Product has Approved by "+uname);
 level++;
 if(level==2)
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
     if(ImageRepository.ImageSet.Forward.Exists()){ 
     ImageRepository.ImageSet.Forward.Click();
     Delay(1000);
   }
}
else{ 
Log.Warning("Product is not Approved by :"+uname);
}

}


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
//      Log.Message(temp)
      }
      else{ 
        temp = temp;
      }
//      }
     Arrays[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}


function CreateProduct(){ 
  STIME = WorkspaceUtils.StartTime();
  ReportUtils.logStep("INFO", "Create Product test started::"+STIME);
  GotoMenu();
  searchcompany()
  Global_client_Data_2();
  WorkspaceUtils.closeAllWorkspaces();
}