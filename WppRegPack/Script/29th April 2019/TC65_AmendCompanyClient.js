//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "AmendClient";
var STIME = "";
var SOX_Array = [];
var SOX_Arrays = [];
var checkmark = false;
var approvers ="";
var Approve_Level = [];
var HRData = [];
var LoginEmp = [];
var y = 0;
var UserPasswd = [];
var ClientNo="";

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

    }

function goToCompanyClient(){ 
      
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
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "")
.SWTObject("Button", "Active Customers");
SOX_Arrays = [];
SOX_Arrays = SOXexcel(sheetName,1);
inactive.Click();
Delay(5000);
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
ClientNo.Keys(SOX_Arrays[1]);
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//ClientName.Keys("Automation Client 19December2018 19:53:38");
ClientName.setText(SOX_Arrays[0]);
Delay(5000);
var cltTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var flag = false
  for(var v=0;v<cltTable.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(temp_user[1]+" "+STIME)){ 
      if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==SOX_Arrays[0]){ 
      flag=true;
      ClientNo = cltTable.getItem(v).getText(2).OleValue.toString().trim();
//      cltTable.Keys("[Enter]");
      break;
    }else{ 
    cltTable.Keys("[Down]");
  }
      
  }
  
if(flag){
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();
Delay(5000);
if(ImageRepository.ImageSet.Forward.Exists()){
ImageRepository.ImageSet.Forward.DblClick();
}
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var refr = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
refr.Refresh();
var CompanyClientBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
CompanyClientBar.Click();
Delay(3000);
SOX_Array = [];
SOX_Array = SOXexcel(sheetName,3);
Delay(2000);
var companyclient  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  companyclient.Click()
  var table  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2)
Delay(3000);
  flag=false;
  Log.Message(SOX_Array[0])
  for(var v=0;v<table.getItemCount();v++){ 
      if((table.getItem(v).getText_2(0).OleValue.toString().trim()==SOX_Array[0])&&(table.getItem(v).getText_2(1).OleValue.toString().trim()==SOX_Array[1])){ 
      flag=true;
      table.Keys("[Down]");
      table.Keys("[Enter]");
      break;
    }else{ 
    table.Keys("[Down]");
//Sys.Desktop.KeyDown(0x28);
//Sys.Desktop.KeyUp(0x28);
  }
      
  }
if(flag){
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();

var infoBar  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
infoBar.Click();
Delay(6000);

    ImageRepository.ImageSet.Maximize.Click();
    Delay(5000);
    
var comapnyName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
comapnyName.Click();
if((SOX_Array[1]!="")&& (SOX_Array[1]!=comapnyName.getText())){
comapnyName.Keys("^a[BS]");
comapnyName.setText(SOX_Array[1] + " "+STIME );
checkmark = true;
ReportUtils.logStep("INFO", "comapnyName has Changed");
}
var street1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
street1.Click();
if((SOX_Array[2]!="")&& (SOX_Array[2]!=street1.getText())){
street1.Keys("^a[BS]");
street1.setText(SOX_Array[2]);
checkmark = true;
ReportUtils.logStep("INFO", "street1 has Changed");
}
var street2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
street2.Click();
if((SOX_Array[3]!="")&& (SOX_Array[3]!=street2.getText())){
street2.Keys("^a[BS]");
street2.setText(SOX_Array[3]);
checkmark = true;
ReportUtils.logStep("INFO", "street2 has Changed");
}
var street3 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
street3.Click();
if((SOX_Array[4]!="")&& (SOX_Array[4]!=street3.getText())){
street3.Keys("^a[BS]");
street3.setText(SOX_Array[4]);
checkmark = true;
ReportUtils.logStep("INFO", "street3 has Changed");
}
var area = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
area.Click();
if((SOX_Array[5]!="")&& (SOX_Array[5]!=area.getText())){
area.Keys("^a[BS]");
area.setText(SOX_Array[5]);
checkmark = true;
ReportUtils.logStep("INFO", "area has Changed");
}
var Attn =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2);
Attn.Click();
if((SOX_Array[18]!="")&& (SOX_Array[18]!=Attn.getText())){
Attn.Keys("^a[BS]");
Attn.setText(SOX_Array[18]);
checkmark = true;
ReportUtils.logStep("INFO", "Attn has Changed");
}
var Email = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 13).SWTObject("McTextWidget", "", 2); 
Email.Click();
if((SOX_Array[19]!="")&& (SOX_Array[19]!=Email.getText())){
Email.Keys("^a[BS]");
Email.setText(SOX_Array[19]);
checkmark = true;
ReportUtils.logStep("INFO", "Email has Changed");
}
var phone =  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
phone.Click();
if((SOX_Array[20]!="")&& (SOX_Array[20]!=phone.getText())){
phone.Keys("^a[BS]");
phone.setText(SOX_Array[20]);
checkmark = true;
ReportUtils.logStep("INFO", "phone has Changed");
}
var Fax = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2); 
Fax.Click();
if((SOX_Array[21]!="")&& (SOX_Array[21]!=Fax.getText())){
Fax.Keys("^a[BS]");
Fax.setText(SOX_Array[21]);
checkmark = true;
ReportUtils.logStep("INFO", "Fax has Changed");
}
var postalcode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
postalcode.Click();
if((SOX_Array[6]!="")&& (SOX_Array[6]!=postalcode.getText())){
postalcode.Keys("^a[BS]");
postalcode.setText(SOX_Array[6]);
checkmark = true;
ReportUtils.logStep("INFO", "postalcode has Changed");
}
var postalDistrict = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
postalDistrict.Click();
if((SOX_Array[7]!="")&& (SOX_Array[7]!=postalDistrict.getText())){
postalDistrict.Keys("^a[BS]");
postalDistrict.setText(SOX_Array[7]);
checkmark = true;
ReportUtils.logStep("INFO", "postalDistrict has Changed");
}
var AccountManger = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 15).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3); 
if((SOX_Array[23]!="")&& (SOX_Array[23]!=AccountManger.getText())){
AccountManger.Click();
AccountManger.Keys("^a[BS]");
var checkmarks = false;
//checkmarks = WorkspaceUtils.SearchByValue(AccountManger,"Employee",SOX_Array[23]);

  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(SOX_Array[23]);
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
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==SOX_Array[23]) { 
        Sys.Desktop.KeyDown(0x28); // Down Arrow
        Delay(1000);
        Sys.Desktop.KeyUp(0x28); 
        Sys.Desktop.KeyDown(0x0D);
        Sys.Desktop.KeyUp(0x0D);
        checkmark = true;
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          AccountManger.setText("");
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
          AccountManger.setText("");
    }
    
if(checkmarks){
ReportUtils.logStep("INFO", "AccountManger ID has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"AccountManger ID is NOT Available")
}
}
var country  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
if((SOX_Array[8]!="")&& (SOX_Array[8]!=country.getText())){
country.Keys("United Kingdom");
Delay(5000);
country.Click();
//Sys.Process("Maconomy").Refresh();
var checkmarks = false;
checkmarks = DropDownList(SOX_Array[8])
if(checkmarks){
ReportUtils.logStep("INFO", "Country has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"Country is NOT Available")
}
}
var invoiceSpec = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2); 
if((SOX_Array[17]!="")&& (SOX_Array[17]!=invoiceSpec.getText())){
invoiceSpec.Click();
invoiceSpec.Keys(SOX_Array[17]);
Delay(5000);
checkmark = true;
ReportUtils.logStep("INFO", "invoiceSpec has Changed");
}
var currency = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 14).SWTObject("McPopupPickerWidget", "", 2);
if((SOX_Array[12]!="")&& (SOX_Array[12]!=currency.getText())){
currency.Keys("GBP");
Delay(5000);
currency.Click();
//Sys.Process("Maconomy").Refresh();
var checkmarks = false;
checkmarks = DropDownList(SOX_Array[12])
if(checkmarks){
ReportUtils.logStep("INFO", "currency has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"currency is NOT Available")
}
}
var partyBFC = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 3);
if((SOX_Array[14]!="")&& (SOX_Array[14]!=partyBFC.getText())){
partyBFC.Click();
partyBFC.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(partyBFC,"Counter Party BFC",SOX_Array[14]);
if(checkmarks){
ReportUtils.logStep("INFO", "partyBFC has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"partyBFC is NOT Available")
}
}
var JobPriceListCost = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if((SOX_Array[25]!="")&& (SOX_Array[25]!=JobPriceListCost.getText())){
JobPriceListCost.Click();
JobPriceListCost.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(JobPriceListCost,"Job Price List",SOX_Array[25]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "JobPriceListCost has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"JobPriceListCost is NOT Available")
}
}
var JobPriceListInterCom = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
if((SOX_Array[26]!="")&& (SOX_Array[26]!=JobPriceListInterCom.getText())){
JobPriceListInterCom.Click();
JobPriceListInterCom.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(JobPriceListInterCom,"Job Price List",SOX_Array[26]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "JobPriceListInterCom has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"JobPriceListInterCom is NOT Available")
}
}
var JobPriceListSale = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
if((SOX_Array[27]!="")&& (SOX_Array[27]!=JobPriceListSale.getText())){
JobPriceListSale.Click();
JobPriceListSale.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(JobPriceListSale,"Job Price List",SOX_Array[27]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "JobPriceListSale has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"JobPriceListSale is NOT Available")
}
}
var JobPriceListStdSale = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
if((SOX_Array[28]!="")&& (SOX_Array[28]!=JobPriceListStdSale.getText())){
JobPriceListStdSale.Click();
JobPriceListStdSale.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(JobPriceListStdSale,"Job Price List",SOX_Array[28]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "JobPriceListStdSale has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"JobPriceListStdSale is NOT Available")
}
}
var JobSurchargeRule = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 6).SWTObject("McValuePickerWidget", "", 2);
if((SOX_Array[29]!="")&& (SOX_Array[29]!=JobSurchargeRule.getText())){
JobSurchargeRule.Click();
JobSurchargeRule.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(JobSurchargeRule,"Job Surcharge Rule",SOX_Array[29]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "JobSurchargeRule has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"JobSurchargeRule is NOT Available")
}
}






adr = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");
adr.Click();
adr.MouseWheel(-200);









var language = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
if((SOX_Array[9]!="")&& (SOX_Array[9]!=language.getText())){
language.Keys("English");
Delay(5000);
language.Click();
//Sys.Process("Maconomy").Refresh();
var checkmarks = false;
checkmarks = DropDownList(SOX_Array[9])
if(checkmarks){
ReportUtils.logStep("INFO", "language has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"language is NOT Available")
}
}
var taxNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2) 
taxNo.Click();
if((SOX_Array[10]!="")&& (SOX_Array[10]!=taxNo.getText())){
taxNo.Keys("^a[BS]");
taxNo.setText(SOX_Array[10]);
checkmark = true;
ReportUtils.logStep("INFO", "taxNo has Changed");
}
var companyRegNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
companyRegNo.Click();
if((SOX_Array[11]!="")&& (SOX_Array[11]!=companyRegNo.getText())){
companyRegNo.Keys("^a[BS]");
companyRegNo.setText(SOX_Array[11]);
checkmark = true;
ReportUtils.logStep("INFO", "companyRegNo has Changed");
}

var clientgroup = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPopupPickerWidget", "", 2)
if((SOX_Array[13]!="")&& (SOX_Array[13]!=clientgroup.getText())){
clientgroup.Keys("Aerospace");
Delay(5000);
clientgroup.Click();
//Sys.Process("Maconomy").Refresh();
var checkmarks = false;
checkmarks = DropDownList(SOX_Array[13])
if(checkmarks){
ReportUtils.logStep("INFO", "clientgroup has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"clientgroup is NOT Available")
}
}

var modacode = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 5).SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
if((SOX_Array[15]!="")&& (SOX_Array[15]!=modacode.getText())){
modacode.Click();
modacode.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(modacode,"Option",SOX_Array[15]);
if(checkmarks){
ReportUtils.logStep("INFO", "MODA code has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"MODA Code is NOT Available")
}
}
var parentClient = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("McGroupWidget", "", 5).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2); 
if((SOX_Array[16]!="")&& (SOX_Array[16]!=parentClient.getText())){
parentClient.Click();
parentClient.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(parentClient,"Option",SOX_Array[16]);
if(checkmarks){
ReportUtils.logStep("INFO", "parentClient has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"parentClient is NOT Available")
}
}


var AccountDirector = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 13).SWTObject("McValuePickerWidget", "", 2); 
if((SOX_Array[22]!="")&& (SOX_Array[22]!=AccountDirector.getText())){
AccountDirector.Click();
AccountDirector.Keys("^a[BS]");
var checkmarks = false;
//checkmarks = WorkspaceUtils.SearchByValue(AccountDirector,"Employee",SOX_Array[22]);

  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(SOX_Array[22]);
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
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==SOX_Array[22]) { 
        Sys.Desktop.KeyDown(0x28); // Down Arrow
        Delay(1000);
        Sys.Desktop.KeyUp(0x28); 
        Sys.Desktop.KeyDown(0x0D);
        Sys.Desktop.KeyUp(0x0D);
        checkmark = true;
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          AccountDirector.setText("");
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
          AccountDirector.setText("");
    }

if(checkmarks){
ReportUtils.logStep("INFO", "AccountDirector ID has Changed");
checkmark = true;
}
else{ 
  ValidationUtils.verify(checkmarks,true,"AccountDirector ID is NOT Available")
}
}

var companyTaxvalue = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPopupPickerWidget", "", 2); 
if((SOX_Array[24]!="")&& (SOX_Array[24]!=companyTaxvalue.getText())){
companyTaxvalue.Keys("Local");
Delay(5000);
companyTaxvalue.Click();
//Sys.Process("Maconomy").Refresh();
var checkmarks = false;
checkmarks = DropDownList(SOX_Array[24])
if(checkmarks){
ReportUtils.logStep("INFO", "companyTaxvalue has Changed");
checkmark = true;
}
else{ 
  ValidationUtils.verify(checkmarks,true,"companyTaxvalue is NOT Available")
}
}


var BudgetHolder = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if((SOX_Array[30]!="")&& (SOX_Array[30]!=BudgetHolder.getText())){
BudgetHolder.Click();
BudgetHolder.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(BudgetHolder,"Employee",SOX_Array[25]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "BudgetHolder has Changed");
checkmark = true;
}
else{ 
  ValidationUtils.verify(checkmarks,true,"BudgetHolder is NOT Available")
}
}
var MainBiller = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3); 
if((SOX_Array[31]!="")&& (SOX_Array[31]!=MainBiller.getText())){
MainBiller.Click();
MainBiller.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(MainBiller,"Employee",SOX_Array[31]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "MainBiller has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"MainBiller is NOT Available")
}
}
var ClientFinance = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);   
if((SOX_Array[32]!="")&& (SOX_Array[32]!=ClientFinance.getText())){
ClientFinance.Click();
ClientFinance.Keys("^a[BS]");
var checkmarks = false;
checkmarks = WorkspaceUtils.SearchByValue(ClientFinance,"Employee",SOX_Array[32]);
//checkmark = true;
if(checkmarks){
ReportUtils.logStep("INFO", "ClientFinance has Changed");
checkmark = true;
}else{ 
  ValidationUtils.verify(checkmarks,true,"ClientFinance is NOT Available")
}
}



if(checkmark){
var save  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
save.Click();
Delay(4000);


if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption == "Client Management - Information"){
var CIOk = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
CIOk.Click();
}
Delay(2000);
if(ImageRepository.ImageSet.OK_Button.Exists()){ 
  ImageRepository.ImageSet.OK_Button.Click();
}
Delay(3000);
var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
submit.Click();

Delay(4000);
ValidationUtils.verify(true,true,"Company Client has been Changed");
if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
  Delay(3000);
}

  var document = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  document.forceFocus();
  document.setVisible(true);
  document.Click();
    Delay(3000);

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
info.Click();
Delay(5000);
var indiaBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
indiaBar.Click();
Delay(2000);
if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
Delay(3000);
}
var AllApproved = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 9);
AllApproved.Click();
Delay(4000);
y =0 ;
//var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
SOX_Arrays = [];
SOX_Arrays = SOXexcel(sheetName,1);
for(var i=0;i<ApproverTable.getItemCount();i++){ 
  
    approvers="";
    if(ApproverTable.getItem(i).getText_2(8)!="Approved"){
       approvers = SOX_Array[0]+"*"+SOX_Arrays[0]+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +i+ ": " +approvers);
//       Approve_Level[y] = Company_ID+"*"+Job_Name+"*"+approvers;
        Approve_Level[y] = approvers;
       y++;
       }
}
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "").Click();
Delay(3000);
if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
}

HRData = goToHR();
LoginEmp = Credentiallogin(Project.Path+excelName, "userRoles");

if(SOX_Array[0]!="")
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,SOX_Array[0]);
else
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,"");

RestMaconomy(UserPasswd)
}


    }
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
//UserPasswd[0] = "1702*Automation Client 19December2018 19:53:38*1702 - Finance*CORE@WPP123";; **Opco - Finance*1707 - Senior Finance
//UserPasswd[1] = "122219*Regular Hindustan*somsubhra.banerjee@jwt.com*CORE@WPP123";
//UserPasswd[0] = "1706*Automation Client 19December2018 19:53:38*SSC IN -  CT Clients*CORE@WPP123";
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
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "")
.SWTObject("Button", "Active Customers");

inactive.Click();
Delay(5000);
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
//ClientNo.setText(temp_user[0]);
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//ClientName.Keys("Automation Client 19December2018 19:53:38");
ClientName.setText(temp_user[1]);
//ClientName.Keys("[Tab]");
//var AppList = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 3);
//AppList.Click();
//Delay(4000);
//ImageRepository.ImageSet.sale_dropDown.Click();
//Delay(1000);
////Sys.Process("Maconomy").Refresh();
//var code = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);
//    code.Keys("Yes");
//    Delay(5000);
//    Sys.Desktop.KeyDown(0x0D);
//    Sys.Desktop.KeyUp(0x0D);
Delay(5000);
var cltTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var flag = false
  for(var v=0;v<cltTable.getItemCount();v++){ 
//    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(temp_user[1]+" "+STIME)){ 
      if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==temp_user[1]){ 
      flag=true;
      ClientNo = cltTable.getItem(v).getText(2).OleValue.toString().trim();
//      cltTable.Keys("[Enter]");
      break;
    }else{ 
    cltTable.Keys("[Down]");
  }
      
  }
  
if(flag){
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();
Delay(5000);
if(ImageRepository.ImageSet.Forward.Exists()){
ImageRepository.ImageSet.Forward.DblClick();
}
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var refr = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
refr.Refresh();
var CompanyClientBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
CompanyClientBar.Click();
Delay(3000);

Delay(2000);
var companyclient  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  companyclient.Click()
  var table  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2)
Delay(3000);
  flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==temp_user[0]){ 
      flag=true;
      table.Keys("[Down]");
      table.Keys("[Enter]");
      break;
    }else{ 
    table.Keys("[Down]");
//Sys.Desktop.KeyDown(0x28);
//Sys.Desktop.KeyUp(0x28);
  }
      
  }
if(flag){
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();

var infoBar  = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
infoBar.Click();
Delay(6000);

    ImageRepository.ImageSet.Maximize.Click();
    Delay(5000);
    
    var document = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  document.forceFocus();
  document.setVisible(true);
  document.Click();
  
  Delay(3000);
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
info.Click();
Delay(5000);
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve);
if(Approve.isEnabled()){
Approve.Click();
Delay(5000);
ReportUtils.logStep("INFO","Approve Button is Clicked");
ValidationUtils.verify(true,true,"Approved by :"+temp_user[2]);

}
else{ 
  Log.Message("Approve Button is Invisible");
}
WorkspaceUtils.closeAllWorkspaces();
    }
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

function ChangeCompanyClient(){ 
  GotoMenu();
  ReportUtils.logStep("INFO","Entering Company Client Details to Change")
  goToCompanyClient();
}


function DropDownList(value){ 
 var checklist = false;
  var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  var Add_Visible4 = true;
  while(Add_Visible4){
  if(list.isEnabled()){
  Add_Visible4 = false;
      for(var i=0;i<4;i++){
      list.Keys("[Up]");
      }
  
      for(var i=0;i<list.getItemCount();i++){ 
        if(list.getItem(i).getText_2(0)!=null){ 
          if(list.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
            list.Keys("[Enter]");
            Delay(5000);
            checklist = true;
            break;
          }else{ 
            list.Keys("[Down]");
          }
          
        }else{ 
          list.Keys("[Down]");
        }
      }
  }
  }
  return checklist;
}