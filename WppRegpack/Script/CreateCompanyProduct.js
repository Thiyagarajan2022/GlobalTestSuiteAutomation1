//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateCompanyProduct";
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
var settlingcompanyvalue,languageValue,attnValue,emailValue,accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue,jobPricelListSalesValue,clientName,ClientNumber,brandName,brandNumber,Currency,Ph_No,Email,C_BFC,SII_Tax,State,GST,PAN,TAN,TIN="";

// New Compnay Product
function CompanyProductCreation(){
  
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
sheetName = "CreateCompanyProduct";
Language = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
//clientName,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product ="";
brandNumber,brandName,settlingcompanyvalue,languageValue,attnValue,emailValue,accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue,jobPricelListSalesValue,clientName,Ph_No,Email,C_BFC,SII_Tax,State,GST,PAN,TAN,TIN="";

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
//EnvParams.Opco = "1712";
getDetails();
gotoMenu(); 
gotoClientSearch();
NewCompanyBrand();
//CompanyBrandTable();
//if(EnvParams.Country.toUpperCase()=="INDIA"){
//Runner.CallMethod("IND_CreateCompnayClient.indiaSpecific",State,GST,PAN,TAN,TIN);
//}
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

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
for(var i=level;i<ApproveInfo.length;i++){
level=i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApproveClient(temp[1],temp[2],i);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}
WorkspaceUtils.closeAllWorkspaces();

}


function FinalApproveClient(ClientNum,Apvr,lvl){ 

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

ValidationUtils.verify(flag,true,"company product is available in Approval List");
TextUtils.writeLog("company product is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
//var Approve = Aliases.CreateCompanyClient.Composite.SingleToolItemControl;
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.Approve;


var Approve = Aliases.CreateCompanyClient.Composite.SingleToolItemControl;
if((Approve.Enabled)&&(Approve.Visible)){
Sys.HighlightObject(Approve)
Log.Message(Approve.FullName);
}else{
Approve = Aliases.Maconomy.Banking.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite4.SingleToolItemControl2
Sys.HighlightObject(Approve)
Log.Message(Approve.FullName);
}
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(8000, "Waiting To Approve");;
ValidationUtils.verify(true,true,"company product is Approved by "+Apvr)
aqUtils.Delay(8000, Indicator.Text);;
TextUtils.writeLog("company product is Approved by "+Apvr);
if(Approve_Level.length==lvl+1){

aqUtils.Delay(8000, Indicator.Text); 

  TextUtils.writeLog("company product Number :"+ClientNum); 
  
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
ValidationUtils.verify(true,false,"company product is not Approved")
      }
}
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("Company Product Number",EnvParams.Opco,"Data Management",ClientNum)
ExcelUtils.WriteExcelSheet("Company Product Name",EnvParams.Opco,"Data Management",ClientName)

  var closeApproval = Aliases.CreateCompanyClient.Composite.PTabItemPanel2.CloseApproverTable;
  Sys.HighlightObject(closeApproval);
 closeApproval.HoverMouse();
 closeApproval.Click();
 ImageRepository.ImageSet.Forward.Click();
 var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
}
  ValidationUtils.verify(true,true,"company product is Approved by "+Apvr)

  
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
  


 ExcelUtils.setExcelName(workBook, "Data Management", true);
  ClientNo = ReadExcelSheet("Company Client Number",EnvParams.Opco,"Data Management");
  ClientNumber =ReadExcelSheet("Company Client Number",EnvParams.Opco,"Data Management");
  if((ClientNo=="")||(ClientNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  ClientNumber =ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  }
  if((ClientNo==null)||(ClientNo=="")){ 
  ValidationUtils.verify(false,true,"Client Number is Needed to Create company product");
  }
    Log.Message("ClientNumber"+ClientNumber)


  
  
ExcelUtils.setExcelName(workBook, sheetName, true);
brandName = ExcelUtils.getRowDatas("Product Name",EnvParams.Opco)
//  }
if((brandName==null)||(brandName=="")){ 
ValidationUtils.verify(false,true,"Product Name is Needed to Create company product");
}



ExcelUtils.setExcelName(workBook, sheetName, true);
brandNumber = ExcelUtils.getRowDatas("Product Number",EnvParams.Opco)

if((brandNumber==null)||(brandNumber=="")){ 
ValidationUtils.verify(false,true,"Product Number is Needed to Create company product");
}
Log.Message("Product Number"+brandNumber)



ExcelUtils.setExcelName(workBook, sheetName, true);

  Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Create a company product");

}
Log.Message("Currency"+Currency)
  
settlingcompanyvalue = ExcelUtils.getRowDatas("Settling company",EnvParams.Opco)
if((settlingcompanyvalue==null)||(settlingcompanyvalue=="")){ 
ValidationUtils.verify(false,true,"settlingcompanyvalue is Needed to Create a company product");
}

languageValue = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((languageValue==null)||(languageValue=="")){ 
ValidationUtils.verify(false,true,"language is Needed to Create a company product");
}

attnValue = ExcelUtils.getRowDatas("Attn",EnvParams.Opco)
if((attnValue==null)||(attnValue=="")){ 
ValidationUtils.verify(false,true,"attnValue is Needed to Create a company product");
}


emailValue = ExcelUtils.getRowDatas("Email",EnvParams.Opco)
if((emailValue==null)||(emailValue=="")){ 
ValidationUtils.verify(false,true,"emailValue is Needed to Create a company product");
}
Log.Message(emailValue);

accountDirectorNoValue = ExcelUtils.getRowDatas("AccountDirectorNo",EnvParams.Opco)
if((accountDirectorNoValue==null)||(accountDirectorNoValue=="")){ 
ValidationUtils.verify(false,true,"accountDirectorNoValue is Needed to Create a company product");
}

controlAccountNoValue = ExcelUtils.getRowDatas("ControlAccountNo",EnvParams.Opco)
if((controlAccountNoValue==null)||(controlAccountNoValue=="")){ 
ValidationUtils.verify(false,true,"controlAccountNoValue is Needed to Create a company product");
}

paymentTermsValue = ExcelUtils.getRowDatas("PaymentTerms",EnvParams.Opco)
if((paymentTermsValue==null)||(paymentTermsValue=="")){ 
ValidationUtils.verify(false,true,"paymentTermsValue is Needed to Create a company product");
}


jobPricelListSalesValue = ExcelUtils.getRowDatas("JobPricelListSales",EnvParams.Opco)
if((jobPricelListSalesValue==null)||(jobPricelListSalesValue=="")){ 
ValidationUtils.verify(false,true,"jobPricelListSalesValue No. is Needed to Create a company product");
}

Log.Message(jobPricelListSalesValue);

companyTaxCodeValue = ExcelUtils.getRowDatas("CompanyTaxCode",EnvParams.Opco)
if((companyTaxCodeValue==null)||(companyTaxCodeValue=="")){ 
ValidationUtils.verify(false,true,"companyTaxCodeValue is Needed to Create a company product");
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
  var GblClient = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.GlobalClientTab;
  
  GblClient.Click();
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 var CompanyNumber = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.CompanyNo;

  CompanyNumber.Click();
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");

//  WorkspaceUtils.SearchByValue(CompanyNumber,"Company",EnvParams.Opco,"Company Number");

 var curr = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Currency
 
//Aliases.CreateCompanyClient.Composite.CurrencyPicker;
 
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
 var ClientNumber = Aliases.CreateCompanyClient.Composite.ClientNoSearch;
  if(ClientNo!=""){
  ClientNumber.Click();
  WorkspaceUtils.VPWSearchByValue(ClientNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNo,"Client Number");
    }
    
 var ClientName =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.ClientName;

 ClientName.setText("*");
 
 
 var save = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.Save;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 
}



function NewCompanyBrand(){ 
  
 var active = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());

  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var table = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.GlobalClientTable;
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
  //  table.getItem(0).
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  
   aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
   aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var Sublevels =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SubLevels;
  waitForObj(Sublevels);
  Sublevels.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var GlobalSublevels =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.GlobalClientTab;
  GlobalSublevels.Click();
   
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var activeProduct = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Products").OleValue.toString().trim());
  Sys.HighlightObject(activeProduct); 
  activeProduct.HoverMouse();
  activeProduct.HoverMouse();
  activeProduct.Click();
  TextUtils.writeLog("Active Global Product is selected");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var productNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(productNmae); 
  Sys.Desktop.Keys("[Up]");
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  productNmae.Keys(brandName);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if((table.getItem(v).getText_2(0).OleValue.toString().trim()==brandNumber)&&(table.getItem(v).getText_2(1).OleValue.toString().trim()==brandName)){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Created Client is available in Approval List");
TextUtils.writeLog("Created Client is available in Approval List");
  
  
  var NewCompanyBrand =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 6);
  waitForObj(NewCompanyBrand);
  NewCompanyBrand.Click();
aqUtils.Delay(8000, Indicator.Text);
    var settlingCompany = Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.SettlingCompany;
    if(settlingcompanyvalue!=""){
   settlingCompany.Click();
  WorkspaceUtils.SearchByValue(settlingCompany,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),settlingcompanyvalue,"Company Number");
  }
  
  var LangaugeDropdown = Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.LanguageDropdown;
   
   if(languageValue!=""){
  LangaugeDropdown.Click();
  WorkspaceUtils.DropDownList(languageValue,"Language")
  }
  
  
 // accountDirectorNoValue,controlAccountNoValue,paymentTermsValue,companyTaxCodeValue
  
  var attn = Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Attn;
   attn.setText(attnValue);
   
   var email =Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.Email;
   
   email.setText(emailValue);
       
  var AccountDirectorNo =  Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.AcDirector;
  if(accountDirectorNoValue!=""){
  AccountDirectorNo.Click();
  WorkspaceUtils.SearchByValue_Emp(AccountDirectorNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),accountDirectorNoValue,"Employee Number");
  }
   
   var controlAccount =Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.ControlAccount;
  if(controlAccountNoValue!=""){
  controlAccount.Click();
  WorkspaceUtils.DropDownList(controlAccountNoValue,"Control Account")
  }
   
  var paymentTerms =Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.PaymentTerms;
  if(paymentTermsValue!=""){
  paymentTerms.Click();
  WorkspaceUtils.DropDownList(paymentTermsValue,"Payment Terms")
  }
   
  var companyTaxCode =Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.CompanyTaxCode;
  if(companyTaxCodeValue!=""){
  companyTaxCode.Click();
  WorkspaceUtils.DropDownList(companyTaxCodeValue,"Company Tax Code")
  }
   
  var JobPricelListSales= Aliases.CreateCompanyBrand_1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.JobPriceListSales;
  if(jobPricelListSalesValue!=""){
   JobPricelListSales.Click();
  
    WorkspaceUtils.SearchByValue(JobPricelListSales,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Price List").OleValue.toString().trim(),jobPricelListSalesValue,"Job Price List Sales");
  }
  
  var CreateBrandButton = Aliases.CreateCompanyBrand_1.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
  Sys.HighlightObject(CreateBrandButton);
  CreateBrandButton.Click();
  
  aqUtils.Delay(5000, Indicator.Text); 
  
  
// if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Global Client - Company Specific Client Information Card")    
//  {
//  var button = Aliases.CreateCompanyBrandPopup.Composite.PopupOK;
//  var label = Sys.Process("Maconomy").SWTObject("Shell", "*")
//
//  Log.Message(label );
//  button.HoverMouse();
//  ReportUtils.logStep_Screenshot("");
//  button.Click();
//  }


var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Company Specific Client Information Card").OleValue.toString().trim()).SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Company Specific Client Information Card").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();

aqUtils.Delay(3000, "Client is Created");
    aqUtils.Delay(10000, Indicator.Text); 
  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Company Specific Client Information Card").OleValue.toString().trim(), 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Company Specific Client Information Card").OleValue.toString().trim()).SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Company Specific Client Information Card").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}


if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var compClientTab =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.CompanySubLevels;
  compClientTab.Click();
  var blockedCompanyTab =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.Active.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blocked").OleValue.toString().trim())
  blockedCompanyTab.Click();  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(2000, Indicator.Text);
  var CompanySublevels =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.CompanySubLevels;
  CompanySublevels.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(2000, Indicator.Text);
    var table =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.ClientSearchTable;
    Sys.HighlightObject(table);
    var Client_Type = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.ClientSearchTable.SWTObject("McPopupPickerWidget", "") ;
   Client_Type.Keys("[Tab][Tab]");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var brandNameField =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.ClientSearchTable.BrandNameField;
    brandNameField.Click();
  brandNameField.Keys(brandName);
  
  aqUtils.Delay(4000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==brandNumber){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"company product is available in maconomy");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==brandNumber){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"company product is available in maconomy");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==brandNumber){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"company product is available in maconomy");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==brandNumber){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"company product is available in maconomy to");
  }
  
  aqUtils.Delay(5000, Indicator.Text);
  
  
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
  
function CompanyBrandTable()
{
      
//  var compClientTab =NameMapping.Sys.Maconomy.CreateCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
//  compClientTab.Click();
  var blockedCompanyTab =Aliases.CreateCompanyClient.Composite.CompanyBlockedRadio
  blockedCompanyTab.Click();
    aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }    
  var table = Aliases.CreateCompanyClient.Composite.CompanyClientTableBlocked;
       


  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==clientName){
  //  table.getItem(0).
  table.HoverMouse(51, 60);
  ReportUtils.logStep_Screenshot();
  table.Click(51, 60);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==clientName){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==clientName){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to block");
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==clientName){
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
  
if(EnvParams.Country.toUpperCase()=="INDIA" || EnvParams.Country.toUpperCase()=="SPAIN" ||  EnvParams.Country.toUpperCase()=="EGYPT" || EnvParams.Country.toUpperCase()=="QATAR" || EnvParams.Country.toUpperCase()=="LEBANON"  || EnvParams.Country.toUpperCase()=="UAE"){
var doc = 
//Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 11);
Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 13)           
}
else{ 
var doc = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 13)

}
  Sys.HighlightObject(doc);
  doc.HoverMouse();
  doc.HoverMouse();
  doc.HoverMouse();
  doc.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var attchDocument =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.AttachDocButton;
  
// Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.Composite.AttachDocument;
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  Sys.HighlightObject(attchDocument);
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
//  ReportUtils.logStep_Screenshot();
var docTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(docTable)
if(docTable.getItemCount()==0){
  attchDocument.Click();
  aqUtils.Delay(4000, "Waiting to Open file");;
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
  aqUtils.Delay(2000, "Waiting to Submit");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  if(Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.isVisible())
var info = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.InfoTab;
else
var info = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl3;
  
//Aliases.CreateCompanyClient.Composite.InfoTAB
  
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Information;
  info.HoverMouse();
  info.HoverMouse();
  info.HoverMouse();
  Sys.HighlightObject(info);
  info.HoverMouse();
  info.HoverMouse();
  info.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    aqUtils.Delay(2000, Indicator.Text);
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var SaveStat = false;
  var phno = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 2)
//  var phno = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.PhoneNo;
  if((phno.getText()=="")||(phno.getText()==null)){
  Sys.HighlightObject(phno)
  WorkspaceUtils.waitForObj(phno)
  phno.Click();
  phno.setText(Ph_No)
  SaveStat = true;
    }
  Delay(3000);
  var Email = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2)
//  var Email = Aliases.Maconomy.AmendCompanyClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Mail;
  if((Email.getText()=="")||(Email.getText()==null)){
  Sys.HighlightObject(Email)
  WorkspaceUtils.waitForObj(Email)
  Email.Click();
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
  if(Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.isVisible())
var submit =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SubmitBrandButton
else
  var submit = NameMapping.Sys.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;

  Sys.HighlightObject(submit)
  // Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Submit;
  Sys.HighlightObject(submit);
  submit.HoverMouse();
  submit.HoverMouse();
  submit.Click();
}
  
  
  function test()
  {
     //
 //
 //
var info = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.InfoTab;
Sys.HighlightObject(info)
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
aqUtils.Delay(5000, "Finding Approvers");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var ClientApproval =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.ClientApprovalTab;
 //Aliases.CreateCompanyClient.Composite.ComapnyClientApprovalTab;
 // Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.ClientApproval;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 var ClientApprovalTab = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.InfoTab;
 //Aliases.CreateCompanyClient.Composite.ClientApprovalTab;
 //Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.ClientApproval_Tab;
 Sys.HighlightObject(ClientApproval);
 ClientApprovalTab.HoverMouse();
 ClientApprovalTab.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
   var ApproverTable = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.ClientApproverTable;
   //Aliases.CreateCompanyClient.Composite.ApproverTable;
   //Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.ApprovarTable;
   var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Approved").OleValue.toString().trim()){
      approvers = EnvParams.Opco+"*"+brandNumber+"*"+ApproverTable.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
//      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
      Approve_Level[y] = approvers;
      y++;
      }
}

TextUtils.writeLog("Finding approvers for Created company product");
var closeCAList = Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.ClientApproverList;
//Aliases.CreateCompanyClient.Composite.ApproverList
//Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel2.CloseApprovalList;
 Sys.HighlightObject(closeCAList);
 closeCAList.HoverMouse();
 closeCAList.Click();
 
ImageRepository.ImageSet.Forward.Click();


CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
//var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
sheetName = "CreateCompanyProduct";
Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
if(OpCo2[2]==Project_manager){
level = 1;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var Approve =Aliases.CreateCompanyBrand.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite1.ApproveButton;
    var ChildCount = 0;
    var Add = [];
//   var Parent = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite;
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
            if((PChild.isVisible()) && (PChild.ChildCount==1)){
//            Log.Message(PChild.FullName)
//         Log.Message(PChild.Child(j).FullName)
         Add[ChildCount] = PChild;
//         Log.Message(Add[ChildCount].FullName)
         ChildCount++;


     }
     }
     
     var Approve = "";
     var pos = 1000;
     for(var ip=0;ip<Add.length;ip++){ 
     if(Add[ip].Height<pos){ 
       pos = Add[ip].Height;
       Log.Message(pos)
       Approve = Add[ip];
     }     
     }
     
     Sys.HighlightObject(Approve)
     Log.Message(Approve.FullName)
     Approved = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
     if(Approved.Visible){ 
     Approve =  Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2);
     }
     else{ 
     Approve = Approve.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1);  
     }
     Sys.HighlightObject(Approve)

 
Log.Message(Approve.FullName)
Sys.HighlightObject(Approve);
 for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
    Approve = Approve.Child(i);
    Sys.HighlightObject(Approve)
    Log.Message(Approve.FullName)
    break;
  }
}

Sys.HighlightObject(Approve)

Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(8000, "Waiting for Approve");;
ValidationUtils.verify(true,true,"company product is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created Budget");
//aqUtils.Delay(8000, Indicator.Text);;
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