//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateCompanyClient";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
//var clientName,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product ="";
var ClientNo = "";
var Language = "";
var Licence_No,Licence_EndDate = "";
var settlingcompanyvalue,languageValue,attnValue,emailValue,accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue,jobPricelListSalesValue,clientName,ClientNumber,Currency,Ph_No,Email,C_BFC,SII_Tax,State,GST,PAN,TAN,TIN="";

function CompanyClientCreation(){
  

TextUtils.writeLog("Company Client Creation Started"); 
Indicator.PushText("waiting for window to open");
Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
  aqUtils.Delay(10000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "Agency Users", true);
var Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco)
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateCompanyClient";
Language = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
//clientName,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product ="";
settlingcompanyvalue,languageValue,attnValue,emailValue,accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue,jobPricelListSalesValue,clientName,Ph_No,Email,C_BFC,SII_Tax,State,GST,PAN,TAN,TIN="";

ClientNumber = "";
Approve_Level = [];

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}

Language = EnvParams.LanChange(Language);
Log.Message("LAN"+Language)
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Company Client Creation started::"+STIME);
TextUtils.writeLog("Execution Started :"+STIME);

getDetails();
gotoMenu(); 
gotoClientSearch();
NewCompanyClient();
CompanyClientTable();
if(EnvParams.Country.toUpperCase()=="INDIA"){
Runner.CallMethod("IND_CreateCompnayClient.indiaSpecific",State,GST,PAN,TAN,TIN);
}
if(EnvParams.Country.toUpperCase()=="SPAIN"){
Runner.CallMethod("SPA_CompanyClient.spainSpecific",SII_Tax);
}
if(EnvParams.Country.toUpperCase()=="UAE"){
Runner.CallMethod("UAE_CompanyClient.UAE_Specific",Licence_No,Licence_EndDate);
}

attachDocument();
Information();
ApprvalInformation();
//CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
level=i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApproveClient(temp[1],temp[2],i);
aqUtils.Delay(8000, Indicator.Text);;
//  Delay(3000);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
}
WorkspaceUtils.closeAllWorkspaces();

}




function FinalApproveClient(ClientNum,Apvr,lvl){ 
//aqUtils.Delay(5000, Indicator.Text);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//aqUtils.Delay(2000, Indicator.Text);
//ImageRepository.ImageSet.Show_Filter.Click();
//}


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var ClientName = "";
var table =Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder;
// Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);

if(Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
var showFilter = Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.ShowFilter;
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.ShowFilterList;
waitForObj(table);
Sys.HighlightObject(showFilter);
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var table = Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
var firstCell = Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.ClientTable.ClientSearch;
waitForObj(firstCell);
Sys.HighlightObject(firstCell);
firstCell.HoverMouse();
firstCell.HoverMouse();
firstCell.Click();
firstCell.setText(EnvParams.Opco);
firstCell.Keys("[Tab][Tab]")

aqUtils.Delay(3000, "Reading Data in table");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var Num = Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(Num);
Num.Click();
Num.setText(ClientNum);
aqUtils.Delay(3000, "Reading Data in table");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var closefilter = Aliases.CreateCompanyClient.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.CloseFilter
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.CloseFilterList;
waitForObj(closefilter);
Sys.HighlightObject(closefilter);
closefilter.HoverMouse();
closefilter.HoverMouse(); 
closefilter.HoverMouse();
closefilter.HoverMouse(); 
//aqUtils.Delay(6000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(2).OleValue.toString().trim()==ClientNum){ 
    ClientName = table.getItem(v).getText_2(3).OleValue.toString().trim()
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Client is available in Approval List");
TextUtils.writeLog("Created Client is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var Approve = Aliases.CreateCompanyClient.Composite.SingleToolItemControl;
if((Approve.Enabled)&&(Approve.Visible)){
Sys.HighlightObject(Approve)
Log.Message(Approve.FullName);
}else{
Approve = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl2
Sys.HighlightObject(Approve)
Log.Message(Approve.FullName);
}

//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.Approve;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(8000, "Waiting To Approve");;
ValidationUtils.verify(true,true,"Global Client is Approved by "+Apvr)
aqUtils.Delay(8000, Indicator.Text);;
TextUtils.writeLog("Global Client is Approved by "+Apvr);
if(Approve_Level.length==lvl+1){
  //***
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//Ok.HoverMouse(); 
//ReportUtils.logStep_Screenshot();
//Ok.Click(); 
//***
aqUtils.Delay(8000, Indicator.Text); 

// for(var j=0;j<12;j++){ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Approve Company Customer by Type"){ 
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Company Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//Ok.HoverMouse(); 
//ReportUtils.logStep_Screenshot();
//Ok.Click(); 
//aqUtils.Delay(8000, Indicator.Text); ;  
//}
// }
 
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//  ExcelUtils.WriteExcelSheet("Global Client",EnvParams.Opco,"Data Management",ClientNum)
  TextUtils.writeLog("Company Client Number :"+ClientNum); 
  
// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
 var ClientApproval = Aliases.CreateCompanyClient.Composite.PTabItemPanel.CompanyClientApproverTab;
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
// }

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

// var ClientApproval = Aliases.CreateCompanyClient.Composite.ComapnyClientApprovalTab;
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(ClientApproval)
 //Aliases.CreateCompanyClient.Composite.PTabItemPanel.CompanyClientApproverTab;
 //Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
   var ApproverTable = Aliases.CreateCompanyClient.Composite.McTableWidget.CompanyClientApproverTable;
   //Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(ApproverTable);
  ReportUtils.logStep_Screenshot();
      for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(true,false,"Created Client is not Approved")
      }
}
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Company Client Number",EnvParams.Opco,"Data Management",ClientNum)
ExcelUtils.WriteExcelSheet("Company Client Name",EnvParams.Opco,"Data Management",ClientName)

  var closeApproval = Aliases.CreateCompanyClient.Composite.PTabItemPanel2.CloseApproverTable;
  Sys.HighlightObject(closeApproval);
 closeApproval.HoverMouse();
 closeApproval.Click();
 ImageRepository.ImageSet.Forward.Click();
 var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
}
  ValidationUtils.verify(true,true,"Global Client is Approved by "+Apvr)

  
}
}

} 



function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Account_Receivable.Exists()){
ImageRepository.ImageSet.Account_Receivable.Click();// GL
}
else if(ImageRepository.ImageSet.Acc_Receivable_1.Exists()){
ImageRepository.ImageSet.Acc_Receivable_1.Click();
}
else{
ImageRepository.ImageSet.Acc_Receivable_2.Click();
}


var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
aqUtils.Delay(3000, Indicator.Text);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}


var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim());
}

}  

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receivable Menu");
TextUtils.writeLog("Entering into Client Management from Accounts Receivable Menu");
}


function getDetails(){ 
  

//    ExcelUtils.setExcelName(workBook, sheetName, true);
// ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
//  ClientNumber=ExcelUtils.getRowDatas("Client Number",EnvParams.Opco);
//  if((ClientNo=="")||(ClientNo==null)){
// ValidationUtils.verify(false,true,"ClientNo is Needed to Create a Client");
//  
//  }
  
  
//  clientName = ExcelUtils.getRowDatas("Client Name",EnvParams.Opco)
//if((clientName==null)||(clientName=="")){ 
//ValidationUtils.verify(false,true,"clientName is Needed to Create a Client");
//}


 ExcelUtils.setExcelName(workBook, "Data Management", true);
//  ClientNo = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
//  ClientNumber =ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
//  if((ClientNo=="")||(ClientNo==null)){
//  ExcelUtils.setExcelName(workBook, sheetName, true);
//  ClientNumber =ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
//  ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
//  }
//  if((ClientNo==null)||(ClientNo=="")){ 
//  ValidationUtils.verify(false,true,"Client Number is Needed to Create Company Brand");
//  }
//    Log.Message("ClientNumber"+ClientNo)
    
//ExcelUtils.setExcelName(workBook, "Data Management", true);
//clientName = ReadExcelSheet("Global Client Name",EnvParams.Opco,"Data Management");
//if((clientName=="")||(clientName==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
clientName = ExcelUtils.getRowDatas("Client Name",EnvParams.Opco)
//}
if((clientName==null)||(clientName=="")){ 
ValidationUtils.verify(false,true,"Client Name is Needed to Create Company Client");
}


      ExcelUtils.setExcelName(workBook, sheetName, true);
  Log.Message("ClientNumber"+ClientNumber)
  Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Create Company Brand");

}
Log.Message("Currency"+Currency)
  
//settlingcompanyvalue,language,attnValue,emailValue,accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue,jobPricelListSalesValue
  
 // ExcelUtils.setExcelName(workBook, sheetName, true);
settlingcompanyvalue = ExcelUtils.getRowDatas("Settling company",EnvParams.Opco)
if((settlingcompanyvalue==null)||(settlingcompanyvalue=="")){ 
ValidationUtils.verify(false,true,"settling companyvalue is Needed to Create a Client");
}

languageValue = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((languageValue==null)||(languageValue=="")){ 
ValidationUtils.verify(false,true,"language is Needed to Create a Client");
}

attnValue = ExcelUtils.getRowDatas("Attn",EnvParams.Opco)
if((attnValue==null)||(attnValue=="")){ 
ValidationUtils.verify(false,true,"attnValue is Needed to Create a Client");
}


emailValue = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
if((emailValue==null)||(emailValue=="")){ 
ValidationUtils.verify(false,true,"emailValue is Needed to Create a Client");
}
Log.Message(emailValue);

accountDirectorNoValue = ExcelUtils.getRowDatas("AccountDirectorNo",EnvParams.Opco)
if((accountDirectorNoValue==null)||(accountDirectorNoValue=="")){ 
ValidationUtils.verify(false,true,"accountDirectorNoValue is Needed to Create a Client");
}

controlAccountNoValue = ExcelUtils.getRowDatas("ControlAccountNo",EnvParams.Opco)
if((controlAccountNoValue==null)||(controlAccountNoValue=="")){ 
ValidationUtils.verify(false,true,"controlAccountNoValue is Needed to Create a Client");
}

paymentTermsValue = ExcelUtils.getRowDatas("PaymentTerms",EnvParams.Opco)
if((paymentTermsValue==null)||(paymentTermsValue=="")){ 
ValidationUtils.verify(false,true,"paymentTermsValue is Needed to Create a Client");
}


jobPricelListSalesValue = ExcelUtils.getRowDatas("JobPricelListSales",EnvParams.Opco)
if((jobPricelListSalesValue==null)||(jobPricelListSalesValue=="")){ 
ValidationUtils.verify(false,true,"jobPricelListSalesValue No. is Needed to Create a Client");
}

Log.Message(jobPricelListSalesValue);

companyTaxCodeValue = ExcelUtils.getRowDatas("CompanyTaxCode",EnvParams.Opco)
if((companyTaxCodeValue==null)||(companyTaxCodeValue=="")){ 
ValidationUtils.verify(false,true,"companyTaxCodeValue is Needed to Create a Client");
}

Ph_No = ExcelUtils.getRowDatas("Phone No",EnvParams.Opco)

Email = ExcelUtils.getRowDatas("Email",EnvParams.Opco)

C_BFC = ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)

if(EnvParams.Country.toUpperCase()=="SPAIN"){
SII_Tax = ExcelUtils.getRowDatas("SII Tax Group",EnvParams.Opco)
if((SII_Tax==null)||(SII_Tax=="")){ 
ValidationUtils.verify(false,true,"SII Tax Group is Needed to Create a Global Brand");
}

}


Licence_No,Licence_EndDate = "";
if(EnvParams.Country.toUpperCase()=="UAE"){
Licence_EndDate = ExcelUtils.getRowDatas("Licence End Date",EnvParams.Opco)
if((Licence_EndDate==null)||(Licence_EndDate=="")){ 
ValidationUtils.verify(false,true,"Licence End Date is Needed to Create a Client");
}

Licence_No = ExcelUtils.getRowDatas("Licence No.",EnvParams.Opco)
if((Licence_No==null)||(Licence_No=="")){ 
ValidationUtils.verify(false,true,"Licence No. is Needed to Create a Client");
}


}
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
var xlDriver = DDT.ExcelDriver(workBook,sheetName,true);
var id =0;
var colsList = [];
var temp ="";
while (!DDT.CurrentDriver.EOF()) {
if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
try{
temp = temp+xlDriver.Value(column).toString().trim();
}
catch(e){
temp = "";
}
break;
}
xlDriver.Next();
}
     
if(temp.indexOf("*")!=-1){
var excelData =  temp.split("*");
}else if(temp.length>0){ 
excelData[0] = temp;
}
     
DDT.CloseDriver(xlDriver.Name);
for(var i=0;i<excelData.length;i++)
return excelData;
}


function gotoClientSearch(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 aqUtils.Delay(8000, Indicator.Text);
  
  var GblClient = Aliases.CreateCompanyClient.Composite.GlobalClient;
  
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.GlobalClient;
  GblClient.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
 var CompanyNumber = Aliases.CreateCompanyClient.Composite.CountryName;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  CompanyNumber.Click();
    var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");

//  WorkspaceUtils.SearchByValue(CompanyNumber,"Company",EnvParams.Opco,"Company Number");

 var curr = Aliases.CreateCompanyClient.Composite.CurrencyPicker;
 
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
// var ClientNumber = Aliases.CreateCompanyClient.Composite.ClientNoSearch;
//  if(ClientNo!=""){
//  ClientNumber.Click();
//  WorkspaceUtils.VPWSearchByValue(ClientNumber,"Client",ClientNo,"Client Number");
//    }
    
 var ClientName =Aliases.CreateCompanyClient.Composite.McTextWidget;

 ClientName.setText(clientName);
 
 
 var save = Aliases.CreateCompanyClient.Composite.SaveBut;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 
}



function NewCompanyClient(){ 
 
  aqUtils.Delay(5000, Indicator.Text);

  aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var NewCompanyClient = Aliases.CreateCompanyClient.Composite.NewCompanyClient;
  if(NewCompanyClient.isEnabled()){
  NewCompanyClient.Click();
  aqUtils.Delay(2000, Indicator.Text);
    }
    else{ 

  }
  
  
  
    var settlingCompany = Aliases.SettlingCompany;
    if(settlingcompanyvalue!=""){
   settlingCompany.Click();
  WorkspaceUtils.SearchByValue(settlingCompany,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),settlingcompanyvalue,"Company Number");
  }
  
  var LangaugeDropdown = Aliases.LnaguageSelector;
   
   if(languageValue!=""){
  LangaugeDropdown.Click();
  WorkspaceUtils.DropDownList(languageValue,"Language")
  }
  
  
 // accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue
  
  var attn =Aliases.McTextWidget;
   attn.setText(attnValue);
   
   var email =Aliases.Email;
   
   email.setText(emailValue);
       
  var AccountDirectorNo =  Aliases.AssociateDir;   
  if(accountDirectorNoValue!=""){
  AccountDirectorNo.Click();
  WorkspaceUtils.SearchByValue_Emp(AccountDirectorNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),accountDirectorNoValue,"Employee Number");
  }
   
   var controlAccount =Aliases.ControlAccount;  
  if(controlAccountNoValue!=""){
  controlAccount.Click();
  WorkspaceUtils.DropDownList(controlAccountNoValue,"Control Account")
  }
   
  var paymentTerms =Aliases.PaymentTerm;
  if(paymentTermsValue!=""){
  paymentTerms.Click();
  WorkspaceUtils.DropDownList(paymentTermsValue,"Payment Terms")
  }
   
  var companyTaxCode =Aliases.CompanyTaxCode;  
  if(companyTaxCodeValue!=""){
  companyTaxCode.Click();
  WorkspaceUtils.DropDownList(companyTaxCodeValue,"Company Tax Code")
  }
   
  var JobPricelListSales= Aliases.pricelistSales 
  if(jobPricelListSalesValue!=""){
   JobPricelListSales.Click();
  
    WorkspaceUtils.SearchByValue(JobPricelListSales,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Price List").OleValue.toString().trim(),jobPricelListSalesValue,"Job Price List Sales");
 // WorkspaceUtils.DropDownList(jobPricelListSalesValue,"Job Price List, Sales")
  }
  
//  var NextButton = Aliases.NextButton;
  var NextButton = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
  NextButton.Click();
  
 aqUtils.Delay(8000, Indicator.Text); 
  var ClientDueDeligencePolicyDropdown = Aliases.DeligencePolicy;
//  var dueDeligenceYes = "Yes"
//    if(dueDeligenceYes!=""){
//  ClientDueDeligencePolicyDropdown.Click();
//  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"By choosing 'Yes', I confirm that I have read and understood the above “Due Diligence” policy and have complied with the above terms in this request.")
//  }
//  
  
  
//  //Expand Window
//   Sys.Desktop.KeyDown(0x12);
//  Sys.Desktop.KeyDown(0x20);
//  Sys.Desktop.KeyUp(0x12);
//  Sys.Desktop.KeyUp(0x20);
//  Sys.Desktop.KeyDown(0x58);
//  Sys.Desktop.KeyUp(0x58);
  
  
   var ClientDueDeligencePolicyDropdown = Aliases.DeligencePolicy;
     ClientDueDeligencePolicyDropdown.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",ClientDueDeligencePolicyDropdown)
// ClientDueDeligencePolicyDropdown.Keys("Yes")
 aqUtils.Delay(5000, Indicator.Text); 
     //Expand Window
   Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);
  aqUtils.Delay(4000, Indicator.Text); 
   var NextButton = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
  NextButton.Click();
//  var NextButtonDeligencepolicy = Aliases.NextButton;
//  NextButtonDeligencepolicy.Click();
  
aqUtils.Delay(4000, Indicator.Text);  
  
    //Expand Window
   Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);
  
  aqUtils.Delay(8000, Indicator.Text); 
  
      var newClientIdentificationInfo = Aliases.NewClientIdInfo;
        newClientIdentificationInfo.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",newClientIdentificationInfo)
//      newClientIdentificationInfo.Keys("Yes")
 aqUtils.Delay(500, Indicator.Text);  
      var checks = Aliases.CheckPerformed;
      checks.setText("YES");
      
  aqUtils.Delay(500, Indicator.Text);     
      var documentNatureofClientBusiness =Aliases.NatureOfClientsBusiness;
      documentNatureofClientBusiness.setText("YES");
  aqUtils.Delay(500, Indicator.Text); 
   
      var verifyCompanyOwners =Aliases.LnaguageSelector;
        verifyCompanyOwners.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",verifyCompanyOwners)
//      verifyCompanyOwners.Keys("Yes")
  aqUtils.Delay(500, Indicator.Text); 
  
      var checksverifyCompanyOwners= Aliases.McTextWidget
      checksverifyCompanyOwners.setText("YES");
  aqUtils.Delay(500, Indicator.Text);     
  
      var foriegnJurisdiction =Aliases.ForiegnJurisdiction
        foriegnJurisdiction.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",foriegnJurisdiction)
//      foriegnJurisdiction.Keys("Yes")
 aqUtils.Delay(500, Indicator.Text);      

      var ForeignJurisdictionChecks = Aliases.foreignJurisdictionChecks
      ForeignJurisdictionChecks.setText("YES");
   aqUtils.Delay(500, Indicator.Text); 
   
      var reputationalIssues = Aliases.CompanyTaxCode;
        reputationalIssues.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",reputationalIssues)
//      reputationalIssues.Keys("Yes")
   aqUtils.Delay(500, Indicator.Text); 
   
      var reputationalChecks = Aliases.ReputationalChecks;
      reputationalChecks.setText("YES");
    aqUtils.Delay(500, Indicator.Text);      
     
      var ConfilictOfinterest =   Aliases.potentialInterest;
        ConfilictOfinterest.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",ConfilictOfinterest)
//      ConfilictOfinterest.Keys("Yes")
    aqUtils.Delay(500, Indicator.Text);    
             
      var ConfilictOfinterestChecks =Aliases.Composite16.cONFLICTcHECKS
      ConfilictOfinterestChecks.setText("YES");
    aqUtils.Delay(500, Indicator.Text);        
      var payForServicesRequested =Aliases.PayForservicesRequested;
        payForServicesRequested.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",payForServicesRequested)
//      payForServicesRequested.Keys("Yes")
    aqUtils.Delay(500, Indicator.Text);   
      var payForServicesRequestedChecks =Aliases.PayForServiceChecks
      payForServicesRequestedChecks.setText("YES");
   
      var documentServices =Aliases.DocumentServices;
      documentServices.setText("YES");
    aqUtils.Delay(10000, Indicator.Text); 
      var Create = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
  waitForObj(Create);
  Create.Click();
//      var CreateClient =Aliases.CreateClient;
//      waitForObj(CreateClient);
//      CreateClient.Click();
       aqUtils.Delay(4000, Indicator.Text); 

  aqUtils.Delay(10000, Indicator.Text); 
//      if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management - Due Diligence Checklist").OleValue.toString().trim())    
//    {
//    var button = Aliases.CompanyRegistrationAlreadyUsED
//       button.HoverMouse();
//     ReportUtils.logStep_Screenshot("");
//      button.Click();
//  }

var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();

aqUtils.Delay(3000, "Client is Created");
    aqUtils.Delay(10000, Indicator.Text); 
  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}

  var compClientTab =Aliases.CreateCompanyClient.Composite.CompanyClientTab;
      compClientTab.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//      var blockedCompanyTab =Aliases.CreateCompanyClient.Composite.CompanyBlockedRadio
//      blockedCompanyTab.Click();  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  }

  
  function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo =Aliases.CreateCompanyClient.Composite.Todos
  // Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(10000, Indicator.Text);

 // Aliases.CreateCompanyClient.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.Refresh
  Aliases.CreateCompanyClient.Composite.Composite42.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl
  
  if(Aliases.CreateCompanyClient.Composite.Composite3.Composite.Composite.Visible){
var refresh = Aliases.CreateCompanyClient.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.Refresh;
Log.Message("true")
}

  if(Aliases.CreateCompanyClient.Composite.Composite42.Composite.Composite.Visible){
var refresh =Aliases.CreateCompanyClient.Composite.Composite42.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl
// Aliases.CreateCompanyClient.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.Refresh;
Log.Message("true")
}
refresh.Click();
aqUtils.Delay(25000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
aqUtils.Delay(25000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
if(Aliases.CreateCompanyClient.Composite.Composite3.Composite.Composite.Visible){
Client_Managt = Aliases.CreateCompanyClient.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
Log.Message("true")
}


if(Aliases.CreateCompanyClient.Composite.Composite42.Composite.Composite.Visible){
Client_Managt = Aliases.CreateCompanyClient.Composite.Composite42.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDosList;
Log.Message("true")
}



var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Company Customer by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Customer by Type from To-Dos List");
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Company Customer by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Customer by Type (Substitute) from To-Dos List");
var listPass = true;   
  }
}  
if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Company Customer").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Customer from To-Dos List");
listPass = false; 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Company Customer (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Customer (Substitute) from To-Dos List");
var listPass = true;   
  }
} 
  }

}
  
function CompanyClientTable()
{
  
aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.CreateCompanyClient.Composite.CompanyClientTableBlocked;
       
Sys.HighlightObject(table);
var Client_type = Aliases.CreateCompanyClient.Composite.CompanyClientTableBlocked.SWTObject("McPopupPickerWidget", "");
Client_type.Keys("[Tab][Tab]");
      
aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
   
if(table.getItem(0).getText_2(3).OleValue.toString().trim()==clientName){
//  table.HoverMouse(51, 60);
//  ReportUtils.logStep_Screenshot();
//  table.Click(51, 60);
ClientNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(1).getText_2(3).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(1).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(2).getText_2(3).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(2).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(3).getText_2(3).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(3).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  
  aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
    }
        
function attachDocument(){ 

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

if((EnvParams.Country.toUpperCase()=="INDIA")&&(EnvParams.Country.toUpperCase()=="SPAIN")&&(EnvParams.Country.toUpperCase()=="UAE")){
//if((EnvParams.Country.toUpperCase()=="SPAIN")&&(EnvParams.Country.toUpperCase()=="UAE")){
//var doc = Aliases.CreateCompanyClient.Composite.AttachDocTab
var doc = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 11);
}
else{ 
var doc = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;

}
  Sys.HighlightObject(doc);
  doc.HoverMouse();
  doc.HoverMouse();
  doc.HoverMouse();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  doc.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var attchDocument =Aliases.CreateCompanyClient.Composite.AttachNewDocument;
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  Sys.HighlightObject(attchDocument);
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  ReportUtils.logStep_Screenshot();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

var docTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(docTable)
if(docTable.getItemCount()==0){
  attchDocument.Click();
  aqUtils.Delay(5000, "Waiting to Open file");;
  var dicratory = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Open file").OleValue.toString().trim(), 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, "Document Attached");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}

}

function Information(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var info = Aliases.CreateCompanyClient.Composite.InfoTAB
//  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab
  info.HoverMouse();
  info.HoverMouse();
  info.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  Sys.HighlightObject(info);
  info.HoverMouse();
  info.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  info.Click();
  aqUtils.Delay(2000, Indicator.Text);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var SaveStat = false;
  var phno = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2)
//  var phno = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.PhoneNo;
  if((phno.getText()=="")||(phno.getText()==null)){
  Sys.HighlightObject(phno)
  WorkspaceUtils.waitForObj(phno)
  phno.setText(Ph_No)
  SaveStat = true;
    }
  Delay(3000);
  var Email = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2)
//  var Email = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Mail;
  if((Email.getText()=="")||(Email.getText()==null)){
  Sys.HighlightObject(Email)
  WorkspaceUtils.waitForObj(Email)
  Email.setText(emailValue)
  SaveStat = true;
    }
    
  var BFC = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 3);
//            Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 3)
  if((BFC.getText()=="")||(BFC.getText()==null)){
  Sys.HighlightObject(BFC)
  WorkspaceUtils.waitForObj(BFC)
  BFC.setText(C_BFC);
  SaveStat = true;
    }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(SaveStat){ 
   var save =  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
//   var save =  Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SaveButton
    save.Click();
  }
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Delay(3000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var submit =Aliases.CreateCompanyClient.Composite.SubmitClientButton;
  // Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Submit;
  Sys.HighlightObject(submit);
  submit.HoverMouse();
  submit.HoverMouse();
  submit.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
  Delay(3000);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
}
    
 
  
  
  function test()
  {
    

var test =Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel;
Sys.HighlightObject(test);

  }



function SearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  waitForObj(code);
  code.Click();
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())


    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();

          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    
    return checkmark;
}


function CredentialLogin(){ 
  var AppvLevl = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }
//WorkspaceUtils.closeAllWorkspaces();
}


function ApprvalInformation(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var ClientApproval =Aliases.CreateCompanyClient.Composite.ComapnyClientApprovalTab;
 // Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.ClientApproval;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
 var ClientApprovalTab = Aliases.CreateCompanyClient.Composite.ClientApprovalTab;
 //Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.ClientApproval_Tab;
 Sys.HighlightObject(ClientApproval);
 ClientApprovalTab.HoverMouse();
 ClientApprovalTab.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
   var ApproverTable = Aliases.CreateCompanyClient.Composite.ApproverTable;
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
   var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Approved").OleValue.toString().trim()){
      approvers = EnvParams.Opco+"*"+ClientNumber+"*"+ApproverTable.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
//      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
      Approve_Level[y] = approvers;
      y++;
      }
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
TextUtils.writeLog("Finding approvers for Created Global Client");
var closeCAList = Aliases.CreateCompanyClient.Composite.ApproverList
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel2.CloseApprovalList;
 Sys.HighlightObject(closeCAList);
 closeCAList.HoverMouse();
 closeCAList.Click();
 
ImageRepository.ImageSet.Forward.Click();


CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
//var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
//var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
//var sheetName = "Agency Users";
//workBook = Project.Path+excelName;
//ExcelUtils.setExcelName(workBook, sheetName, true);
//OpCo2 = ExcelUtils.AgencyLogin(OpCo2,EnvParams.Opco);
sheetName = "CreateCompanyClient";
if(OpCo2[2]==Project_manager){
level = 1;
var Approve = Aliases.CreateCompanyClient.Composite.ApproveButton;
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Approve;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(8000, "Waiting for Approve");;
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created Budget");
aqUtils.Delay(8000, Indicator.Text);;
//  Delay(3000);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
  
}
}
//var Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Approve;
// Sys.HighlightObject(Approve);
// Approve.HoverMouse();
// Approve.Click();
}

function VPWSearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
//    aqUtils.Delay(3000, Indicator.Text);;

    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
  waitForObj(code);
  code.Click();
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();    
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
    var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
    waitForObj(OK);
    Sys.HighlightObject(table); 
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}