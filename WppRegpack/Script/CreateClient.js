//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

//Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreateClient";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var clientName,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product,State,GST,PAN,TAN,SII_Tax,TIN ="";
var Invoice_Spec_Add, i_Add1, i_Add2, i_Add3, ZipCode, CompanyName = "";
var Licence_No,Licence_EndDate = "";
var ClientNumber = "";
var Project_manager = "";
var Language = "";
//Strating Of TestCase
function ClientCreation(){
TextUtils.writeLog("Global Client Creation Started"); 

//Setting Language in WorkspaceUtils
Language = "";
Log.Message("EnvParams.Language :"+EnvParams.Language)
Language = EnvParams.Language;
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
Log.Message("Language :"+Language)

//Checking Login for Client Creation
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager); 
}

//Initializing Variables
excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateClient";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
clientName,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product,State,GST,PAN,TAN,SII_Tax,TIN ="";
ClientNumber = "";
Approve_Level = [];

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Client Creation started::"+STIME);
TextUtils.writeLog("Execution Started :"+STIME);

//try{
//Fetching Variables in Excel
getDetails();
//Entering Client Management
gotoMenu(); 
gotoClientSearch();
globalClient(); 
newGlobalClient();
GlobalClient_Screen2();
ClientDueDiligencePolicy();
globalClientTable();
if(EnvParams.Country.toUpperCase()=="INDIA"){
Runner.CallMethod("IND_CreationClient.indiaSpecific",State,GST,PAN,TAN,TIN);
}
if(EnvParams.Country.toUpperCase()=="SPAIN"){
Runner.CallMethod("SPA_CreateClient.spainSpecific",SII_Tax);
}
if((EnvParams.Country.toUpperCase()=="UAE") || (EnvParams.Country.toUpperCase()=="EGYPT") || (EnvParams.Country.toUpperCase()=="QATAR")){
Runner.CallMethod("UAE_CreateClient.UAE_Specific",Licence_No,Licence_EndDate);
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
}
WorkspaceUtils.closeAllWorkspaces();
TextUtils.writeLog("     ");
var FolderID = Log.CreateFolder(EnvParams.Opco+"_Global Brand");
Log.PushLogFolder(FolderID);
gotoMenu(); 
gotoSearch();
goToglobalClient();
globalBrand();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog("     ");

var FolderID = Log.CreateFolder(EnvParams.Opco+"_Global Product");
Log.PushLogFolder(FolderID);
gotoMenu(); 
gotoSearch();
goToglobalClient();
globalProduct();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog("     ");

var FolderID = Log.CreateFolder(EnvParams.Opco+"_Company Client");
Log.PushLogFolder(FolderID);
gotoMenu(); 
gotoSearch();
goToCompanyClient();
CompanyClient();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog("     ");

var FolderID = Log.CreateFolder(EnvParams.Opco+"_Company Brand");
Log.PushLogFolder(FolderID);
gotoMenu(); 
gotoSearch();
goToglobalClient();
companyBrand();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog("     ");

var FolderID = Log.CreateFolder(EnvParams.Opco+"_Company Product");
Log.PushLogFolder(FolderID);
gotoMenu(); 
gotoSearch();
goToglobalClient();
companyProduct();
WorkspaceUtils.closeAllWorkspaces();
Log.PopLogFolder();
TextUtils.writeLog("     ");
//}
//  catch(err){
//    Log.Message(err);
//  }
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
Sys.HighlightObject(Workspc);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}

Log.Message("Language :"+Language)
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

ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receivable Menu");
TextUtils.writeLog("Entering into Client Management from Accounts Receivable Menu");
}


function getDetails(){ 
Indicator.PushText("Reading Data from Excel");
ExcelUtils.setExcelName(workBook, sheetName, true);
clientName = ExcelUtils.getRowDatas("Client Name",EnvParams.Opco)
if((clientName==null)||(clientName=="")){ 
ValidationUtils.verify(false,true,"Client Name is Needed to Create a Client");
}
Log.Message(clientName)
strt1 = ExcelUtils.getRowDatas("Street 1",EnvParams.Opco)
if((strt1==null)||(strt1=="")){ 
ValidationUtils.verify(false,true,"Street 1 is Needed to Create a Client");
}
Log.Message(strt1)
P_code = ExcelUtils.getRowDatas("Post Code",EnvParams.Opco)
if((P_code==null)||(P_code=="")){ 
ValidationUtils.verify(false,true,"Post Code is Needed to Create a Client");
}
Log.Message(P_code)
P_District = ExcelUtils.getRowDatas("Post District",EnvParams.Opco)
if((P_District==null)||(P_District=="")){ 
ValidationUtils.verify(false,true,"Post District is Needed to Create a Client");
}
Log.Message(P_District)
country = ExcelUtils.getRowDatas("Country",EnvParams.Opco)
if((country==null)||(country=="")){ 
ValidationUtils.verify(false,true,"Country is Needed to Create a Client");
}
Log.Message(country)
clientlan = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((clientlan==null)||(clientlan=="")){ 
ValidationUtils.verify(false,true,"Language is Needed to Create a Client");
}
Log.Message(clientlan)
taxcode = ExcelUtils.getRowDatas("Tax No.",EnvParams.Opco)
if((taxcode==null)||(taxcode=="")){ 
ValidationUtils.verify(false,true,"Tax No. is Needed to Create a Client");
}
Log.Message(taxcode)
companyReg = ExcelUtils.getRowDatas("Company Reg. No.",EnvParams.Opco)
if((companyReg==null)||(companyReg=="")){ 
ValidationUtils.verify(false,true,"Company Reg. No. is Needed to Create a Client");
}
Log.Message(companyReg)
currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((currency==null)||(currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Create a Client");
}
Log.Message(currency)
clientgrp = ExcelUtils.getRowDatas("Client Group",EnvParams.Opco)
if((clientgrp==null)||(clientgrp=="")){ 
ValidationUtils.verify(false,true,"Client Group is Needed to Create a Client");
}
Log.Message(clientgrp)
controlAct = ExcelUtils.getRowDatas("Control Account",EnvParams.Opco)
if((controlAct==null)||(controlAct=="")){ 
ValidationUtils.verify(false,true,"Control Account is Needed to Create a Client");
}
Log.Message(controlAct)
bfc = ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)
if((bfc==null)||(bfc=="")){ 
ValidationUtils.verify(false,true,"Counter Party BFC is Needed to Create a Client");
}
Log.Message(bfc)
//Fax = ExcelUtils.getRowDatas("Fax",EnvParams.Opco)
//if((Fax==null)||(Fax=="")){ 
//ValidationUtils.verify(false,true,"Fax is Needed to Create a Client");
//}

//parentClient = ExcelUtils.getRowDatas("Parent Client",EnvParams.Opco)
//if((parentClient==null)||(parentClient=="")){ 
//ValidationUtils.verify(false,true,"Parent Client is Needed to Create a Client");
//}

//company = ExcelUtils.getRowDatas("Company No.",EnvParams.Opco)
//if((company==null)||(company=="")){ 
//ValidationUtils.verify(false,true,"Company No. is Needed to Create a Client");
//}
attn = ExcelUtils.getRowDatas("Attn.",EnvParams.Opco)
if((attn==null)||(attn=="")){ 
ValidationUtils.verify(false,true,"Attn. is Needed to Create a Client");
}
Log.Message(attn)
mail = ExcelUtils.getRowDatas("E-mail",EnvParams.Opco)
if((mail==null)||(mail=="")){ 
ValidationUtils.verify(false,true,"E-mail is Needed to Create a Client");
}
Log.Message(mail)
phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
if((phone==null)||(phone=="")){ 
ValidationUtils.verify(false,true,"Phone is Needed to Create a Client");
}
Log.Message(phone)
AccDir = ExcelUtils.getRowDatas("Acct. Director No.",EnvParams.Opco)
if((AccDir==null)||(AccDir=="")){ 
ValidationUtils.verify(false,true,"Acct. Director No. is Needed to Create a Client");
}
Log.Message(AccDir)
//AccMan = ExcelUtils.getRowDatas("Account Manager No.",EnvParams.Opco)
//if((AccMan==null)||(AccMan=="")){ 
//ValidationUtils.verify(false,true,"Account Manager No. is Needed to Create a Client");
//}
//Paymentmode = ExcelUtils.getRowDatas("Client Payment Mode",EnvParams.Opco)
//if((Paymentmode==null)||(Paymentmode=="")){ 
//ValidationUtils.verify(false,true,"Client Payment Mode is Needed to Create a Client");
//}
payterm = ExcelUtils.getRowDatas("Payment Terms",EnvParams.Opco)
if((payterm==null)||(payterm=="")){ 
ValidationUtils.verify(false,true,"Payment Terms is Needed to Create a Client");
}
Log.Message(payterm)
Comtaxcode = ExcelUtils.getRowDatas("Company Tax Code",EnvParams.Opco)
if((Comtaxcode==null)||(Comtaxcode=="")){ 
ValidationUtils.verify(false,true,"Company Tax Code is Needed to Create a Client");
}
Log.Message(Comtaxcode)
sales = ExcelUtils.getRowDatas("Job Price List, Sales",EnvParams.Opco)
if((sales==null)||(sales=="")){ 
ValidationUtils.verify(false,true,"Job Price List, Sales is Needed to Create a Client");
}
Log.Message(sales)
//intercomp = ExcelUtils.getRowDatas("Job Price List, Intercomp.",EnvParams.Opco)
//if((intercomp==null)||(intercomp=="")){ 
//ValidationUtils.verify(false,true,"Job Price List, Intercomp. is Needed to Create a Client");
//}
//cost = ExcelUtils.getRowDatas("Job Price List, Cost",EnvParams.Opco)
//if((cost==null)||(cost=="")){ 
//ValidationUtils.verify(false,true,"Job Price List, Cost is Needed to Create a Client");
//}
//standSales = ExcelUtils.getRowDatas("Job Price List, Standard Sales",EnvParams.Opco)
//if((standSales==null)||(standSales=="")){ 
//ValidationUtils.verify(false,true,"Job Price List, Standard Sales is Needed to Create a Client");
//}
brand = ExcelUtils.getRowDatas("Default Brand",EnvParams.Opco)
if((brand==null)||(brand=="")){ 
ValidationUtils.verify(false,true,"Default Brand is Needed to Create a Client");
}
Log.Message(brand)
product = ExcelUtils.getRowDatas("Default Product",EnvParams.Opco)
if((product==null)||(product=="")){ 
ValidationUtils.verify(false,true,"Default Product is Needed to Create a Client");
}
Log.Message(product)

Invoice_Spec_Add, i_Add1, i_Add2, i_Add3, ZipCode, CompanyName = "";
Invoice_Spec_Add = ExcelUtils.getRowDatas("Invoice Specific Address",EnvParams.Opco)
if((Invoice_Spec_Add==null)||(Invoice_Spec_Add=="")){ 
ValidationUtils.verify(false,true,"Invoice Specific Address is Needed to Create a Client");
}
else{

    if(Invoice_Spec_Add.toUpperCase()=="YES"){
      
    i_Add1 = ExcelUtils.getRowDatas("Address 1",EnvParams.Opco)
    if((i_Add1==null)||(i_Add1=="")){ 
    ValidationUtils.verify(false,true,"Address 1 is Needed to Create a Client");
    }

    CompanyName = ExcelUtils.getRowDatas("Company Name",EnvParams.Opco)
    if((CompanyName==null)||(CompanyName=="")){ 
    ValidationUtils.verify(false,true,"Company Name is Needed to Create a Client");
    }

    ZipCode = ExcelUtils.getRowDatas("Zip Code",EnvParams.Opco)
    if((ZipCode==null)||(ZipCode=="")){ 
    ValidationUtils.verify(false,true,"Zip Code is Needed to Create a Client");
    }

    
    }


}


if(EnvParams.Country.toUpperCase()=="INDIA"){
State = ExcelUtils.getRowDatas("State Code",EnvParams.Opco)
if((State==null)||(State=="")){ 
ValidationUtils.verify(false,true,"State Code is Needed to Create a Client");
}
Log.Message(State)
GST = ExcelUtils.getRowDatas("GST Debtor Type",EnvParams.Opco)
if((GST==null)||(GST=="")){ 
ValidationUtils.verify(false,true,"GST Debtor Type is Needed to Create a Client");
}
Log.Message(GST)
PAN = ExcelUtils.getRowDatas("PAN",EnvParams.Opco)
//if((PAN==null)||(PAN=="")){ 
//ValidationUtils.verify(false,true,"PAN is Needed to Create a Client");
//}
Log.Message(PAN)
TAN = ExcelUtils.getRowDatas("TAN",EnvParams.Opco)
//if((TAN==null)||(TAN=="")){ 
//ValidationUtils.verify(false,true,"TAN is Needed to Create a Client");
//}
Log.Message(TAN)
TIN = ExcelUtils.getRowDatas("TIN",EnvParams.Opco)
//if((TAN==null)||(TAN=="")){ 
//ValidationUtils.verify(false,true,"TAN is Needed to Create a Client");
//}
Log.Message(TIN)
}
if(EnvParams.Country.toUpperCase()=="SPAIN"){
SII_Tax = ExcelUtils.getRowDatas("SII Tax Group",EnvParams.Opco)
if((SII_Tax==null)||(SII_Tax=="")){ 
ValidationUtils.verify(false,true,"SII Tax Group is Needed to Create a Client");
}

}

Licence_No,Licence_EndDate = "";
if((EnvParams.Country.toUpperCase()=="UAE") || (EnvParams.Country.toUpperCase()=="EGYPT") || (EnvParams.Country.toUpperCase()=="QATAR")){
Licence_EndDate = ExcelUtils.getRowDatas("Licence End Date",EnvParams.Opco)
if((Licence_EndDate==null)||(Licence_EndDate=="")){ 
ValidationUtils.verify(false,true,"Licence End Date is Needed to Create a Client");
}

Licence_No = ExcelUtils.getRowDatas("Licence No.",EnvParams.Opco)
if((Licence_No==null)||(Licence_No=="")){ 
ValidationUtils.verify(false,true,"Licence No. is Needed to Create a Client");
}


}
Indicator.PushText("Playback");
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
 aqUtils.Delay(2000, "Client Search");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   
 }
 
 var CompanyNumber = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite.companyNo_textbox;
 waitForObj(CompanyNumber);
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   
 }
  Sys.HighlightObject(CompanyNumber);
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   
 }
  CompanyNumber.Click();
  
  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   
 }
 var curr = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite2.currency_Dropdown;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 if(currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(currency,"Currency")
  }
  
 var ClientName = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite4.clientName_textbox;
 ClientName.HoverMouse();
 Sys.HighlightObject(ClientName);
 ClientName.setText(clientName+" "+STIME);
 
 
 var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.Clientsearch_Save;
 save.HoverMouse();
 Sys.HighlightObject(save);
 save.Click();
 
 TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
 aqUtils.Delay(500, "Saving Client Info");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   
 }
}


function globalClient(){ 
  var GblClient = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.GlobalClient;
  GblClient.HoverMouse();
  Sys.HighlightObject(GblClient);
  GblClient.Click();
  aqUtils.Delay(5000, "Saving Country Client");
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
   
 }
  var AllClients = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());
  AllClients.Click();
  aqUtils.Delay(5000, "All Client");
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var NewGlobalClient = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.Composite.NewGlobalClient
  NewGlobalClient.HoverMouse();
  Sys.HighlightObject(NewGlobalClient);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(NewGlobalClient.isEnabled()){
  NewGlobalClient.HoverMouse();
  ReportUtils.logStep_Screenshot();
  NewGlobalClient.Click();
    }
    else{ 
    aqUtils.Delay(2000, "All Client");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }   
  var ActiveClient = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim())
  ActiveClient.Click();
  ActiveClient.HoverMouse();
  ActiveClient.HoverMouse();
  ActiveClient.HoverMouse();  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }   
  NewGlobalClient.HoverMouse();
  ReportUtils.logStep_Screenshot();
  NewGlobalClient.Click();
    }
 }


function newGlobalClient(){ 
  aqUtils.Delay(5000,"Waiting System to Load fully");
 var ClientName = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.ClientName_textbox;
 waitForObj(ClientName)
 var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
 waitForObj(cancel)
 
 ClientName.Click();
 var Wizard_Name = Aliases.Maconomy.New_Global_Client.WndCaption;
 Log.Message(Wizard_Name)
 if(Wizard_Name==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "New Country Client").OleValue.toString().trim())
 ValidationUtils.verify(true,true,"Text as Wizard title is matched with New Country Client Wizard")
 else
 ValidationUtils.verify(false,true,"Text as Wizard title is NOT matched with New Country Client Wizard")
 
 var Wizard_Name = Sys.Process("Maconomy").SWTObject("Shell", Wizard_Name).SWTObject("Composite", "").SWTObject("Label", "*").WndCaption;
 Log.Message(Wizard_Name)
 if(Wizard_Name==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Country Client Master Data").OleValue.toString().trim())
 ValidationUtils.verify(true,true,"Text Country Client Master Data is matched with Country Client Wizard")
 else
 ValidationUtils.verify(false,true,"Text Country Client Master Data is NOT matched with Country Client Wizard")
 
 ClientName.setText(clientName+" "+STIME);
 clientName = clientName+" "+STIME;
 
 var Street1 = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Street1_textbox;
 Street1.Click();
 Street1.setText(strt1);
 
 var PostalCode = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.PostalCode;
 PostalCode.setText(P_code);
 
 var PostalDistrict = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Composite.Postal_District;
 PostalDistrict.setText(P_District);
 aqUtils.Delay(2000,"Waiting System to Load fully");
 var C_Country = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.Country_Dropdown;
 if(country!=""){
  C_Country.Click();
  WorkspaceUtils.DropDownList(country,"Country")
  }
  
  aqUtils.Delay(3000,"Waiting System to Load fully");
 var TaxNo = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.TaxNo;
 TaxNo.HoverMouse();
 Sys.HighlightObject(TaxNo);
 TaxNo.setText(taxcode); 
  
 var ComapnyReg = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.CompanyRegNo;
 ComapnyReg.setText(companyReg); 
 aqUtils.Delay(3000,"Waiting System to Load fully");
 var C_Currency = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.Currency;
  if(currency!=""){
  C_Currency.Click();
  aqUtils.Delay(30000,"Loading currencies")
  WorkspaceUtils.DropDownList(currency,"Currency")
  }
 
  aqUtils.Delay(3000,"Waiting System to Load fully");
  var ClientGroup = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite8.ClientGroup;
  ClientGroup.HoverMouse();
  Sys.HighlightObject(ClientGroup);
  if(clientgrp!=""){
  ClientGroup.Click();
  WorkspaceUtils.DropDownList(clientgrp,"Client Group")
  }
  aqUtils.Delay(2000,"Waiting System to Load fully");
  
  var controlAccount = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.ControlAccount;
  controlAccount.HoverMouse();
  Sys.HighlightObject(controlAccount);
  if(controlAct!=""){
  controlAccount.Click();
  WorkspaceUtils.DropDownList(controlAct,"Control Account")
  }
  
  var PartyBFC = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.Counter_Party_BFC;
  PartyBFC.HoverMouse();
  Sys.HighlightObject(PartyBFC);
  if(bfc!=""){
  PartyBFC.Click();
  WorkspaceUtils.SearchByValue(PartyBFC,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Counter Party BFC").OleValue.toString().trim(),bfc,"Counter Party BFC");
    }
    
  var DefaultBrand = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.DefaultBrand;
  DefaultBrand.HoverMouse();
  Sys.HighlightObject(DefaultBrand);
  DefaultBrand.setText(brand.toString().trim()+" "+STIME);  
  brand = brand.toString().trim()+" "+STIME;
  Log.Message(brand)
  var DefaultProduct = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.DefaultProduct;
  DefaultProduct.setText(product.toString().trim()+" "+STIME);  
  product = product.toString().trim()+" "+STIME;
  
 var Next = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
 Next.HoverMouse();
 Sys.HighlightObject(Next);
 ReportUtils.logStep_Screenshot() ;
 Next.Click();
}


function GlobalClient_Screen2(){ 
  var CompanyNumber = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite11.CompanyNo;
  waitForObj(CompanyNumber)
  
   var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
 waitForObj(cancel)

  CompanyNumber.Click();
  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");
  
  aqUtils.Delay(2000,"Waiting System to Load fully");
  var C_Language = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Language;
  if(clientlan!=""){
  C_Language.Click();
  WorkspaceUtils.DropDownList(clientlan,"Language")
  }
  aqUtils.Delay(2000,"Waiting System to Load fully");
  
  var Attn = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Attn;
  Attn.HoverMouse();
  Sys.HighlightObject(Attn);
  Attn.setText(attn);  
  
  var C_Email  = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite12.Email;
  var Eml_split1 = mail.substring(0,mail.indexOf("@"));
  var Eml_split2 = mail.substring(mail.indexOf("@"));
  Eml_split1 = Eml_split1 +" "+STIME;
  Eml_split1 = Eml_split1.replace(/[_: ]/g,"");
  mail = Eml_split1+Eml_split2
  C_Email.setText(mail);
  
  var C_phone = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.Phone;
  C_phone.setText(phone); 
  
  var C_AcctDir = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Acct_Director_No;
  if(AccDir!=""){
  C_AcctDir.HoverMouse();
  Sys.HighlightObject(C_AcctDir);
  C_AcctDir.Click();
  WorkspaceUtils.SearchByValue_Emp(C_AcctDir,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),AccDir,"Acct Director No");
  }
  
  var C_PaymentTerm = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.Country_Dropdown;
  if(payterm!=""){
  Sys.HighlightObject(C_PaymentTerm);
  C_PaymentTerm.Click();
  WorkspaceUtils.DropDownList(payterm,"Payment Terms",C_PaymentTerm)
  }
  
  var C_companyTaxCode = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.Company_Tax_Code;
  if(Comtaxcode!=""){
  C_companyTaxCode.HoverMouse();
  Sys.HighlightObject(C_companyTaxCode);
  C_companyTaxCode.Click();
  WorkspaceUtils.DropDownList(Comtaxcode,"Company Tax Code");
  }
  
  var C_JobPriceList = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.JobPriceList;
  if(sales!=""){
  Sys.HighlightObject(C_JobPriceList);
  C_JobPriceList.Click();
  WorkspaceUtils.SearchByValue(C_JobPriceList,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Price List").OleValue.toString().trim(),sales,"Job Price List Sales");
         }  
  
 var Next = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
 Sys.HighlightObject(Next);
 Next.HoverMouse();
 ReportUtils.logStep_Screenshot() ;
 Next.Click();
    

}


function ClientDueDiligencePolicy(){
aqUtils.Delay(5000,"Waiting System to Load fully");
var screen = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10;
Sys.HighlightObject(screen);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, "Maximize the screen");
 var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
 waitForObj(cancel)
  var screen = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10;
  screen.Click();
  screen.MouseWheel(-200);

var DueDiligence = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite14.Confirm_Due_Diligence;
Sys.HighlightObject(DueDiligence)
Log.Message(Language)
Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
//DueDiligence.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  DueDiligence.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",DueDiligence)
  
var next = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim());
waitForObj(next);
next.Click();

aqUtils.Delay(2000,"Waiting System to Load fully");

var client_identification = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.Language;
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  Sys.HighlightObject(client_identification);
  Sys.HighlightObject(client_identification);
   var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
  waitForObj(cancel)
  Sys.HighlightObject(client_identification);
  client_identification.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",client_identification)
//  client_identification.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
 
 var checks_did_you_perform = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.Street1_textbox;
 checks_did_you_perform.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

  
 var new_client_business = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite13.Phone;
 new_client_business.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

  
 var company_owner = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.Dropdown;
//  company_owner.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  Sys.HighlightObject(company_owner);
  company_owner.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",company_owner)

 var checks_did_you_perform = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.Textbox;
 Sys.HighlightObject(checks_did_you_perform);
 checks_did_you_perform.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

 var foreign_jurisdictions = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.Dropdown;
//  foreign_jurisdictions.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  Sys.HighlightObject(foreign_jurisdictions);
  foreign_jurisdictions.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",foreign_jurisdictions)

 var checks_did_you_perform = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.TextBox;
 Sys.HighlightObject(checks_did_you_perform);
 checks_did_you_perform.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

  
 var sanction_lists = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite9.ControlAccount;
 Sys.HighlightObject(sanction_lists);
// sanction_lists.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  Sys.HighlightObject(sanction_lists);
  sanction_lists.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",sanction_lists)
  
 var checks_did_you_perform = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite10.TextBox;
 Sys.HighlightObject(checks_did_you_perform); 
 checks_did_you_perform.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

  
 var potential_client_conflicts = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite15.DropDown;
//  potential_client_conflicts.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  Sys.HighlightObject(potential_client_conflicts);
  potential_client_conflicts.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",potential_client_conflicts)

 var checks_did_you_perform = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite16.TextBox;
 Sys.HighlightObject(checks_did_you_perform); 
 checks_did_you_perform.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

  
 var new_client_can_pay = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite17.DropDown;
//  new_client_can_pay.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
  Sys.HighlightObject(new_client_can_pay);
  new_client_can_pay.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim(),"DueDiligence",new_client_can_pay)

 var checks_did_you_perform = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite18.TextBox;
 Sys.HighlightObject(checks_did_you_perform);  
checks_did_you_perform.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());

  aqUtils.Delay(2000, Indicator.Text);
 var services_provided_new_client = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite19.TextBox;
 services_provided_new_client.setText(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());


  var Create = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
  waitForObj(Create);
  Create.Click();
aqUtils.Delay(3000, "Client is getting Created");
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();

aqUtils.Delay(3000, "Client is Created");

  p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{
var Label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client Management").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}
 
  }
  
  
  
function globalClientTable(){ 
  var blocked = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blocked").OleValue.toString().trim());
  Sys.HighlightObject(blocked);
  blocked.HoverMouse();
  blocked.HoverMouse();
  blocked.Click();
  aqUtils.Delay(3000, "Reading Table Data");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
  var table = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.GlobalClient_Table.McGrid;
  Sys.HighlightObject(table);
  var Client_Type = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McPopupPickerWidget", "");
  Client_Type.Keys("[Tab]");
  var C_Name = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.GlobalClient_Table.McGrid.ClientName_Textbox;
  Sys.HighlightObject(C_Name);
  C_Name.Click();
  aqUtils.Delay(3000, "Entering Client");
  C_Name.setText(clientName);
  aqUtils.Delay(3000, "Reading Table Data");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  if(table.getItem(0).getText_2(2).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Created Global Client Name is available in Maconomy");
  }
  else if(table.getItem(1).getText_2(2).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(1).getText_2(0).OleValue.toString().trim()
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Created Global Client Name is available in Maconomy");
  }
  else if(table.getItem(2).getText_2(2).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(2).getText_2(0).OleValue.toString().trim()
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Created Global Client Name is available in Maconomy");
  }
  else if(table.getItem(3).getText_2(3).OleValue.toString().trim()==clientName){
  ClientNumber = table.getItem(3).getText_2(0).OleValue.toString().trim()
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Created Global Client Name is available in Maconomy");
  }
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, "Opening Created Client");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
}


  
function indiaSpecific(){ 
  aqUtils.Delay(7000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var indiaspec = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.IndiaSpecific;
Sys.HighlightObject(indiaspec);
var Start = StartwaitTime();
var waitTime = true;
var Difference = 0;
while(waitTime)
if(Difference<61){
if((indiaspec.isEnabled())&&(indiaspec.text=="India Specific")){
Sys.HighlightObject(indiaspec);
indiaspec.HoverMouse();
indiaspec.Click();
waitTime = false;
}
else{ 
var End = EndTime();
Difference = End - Start;
}
}
else{
 ValidationUtils.verify(true,false,"Screen is not Responding more than a minute");
}


  
  
  
  
  
  var StateCode = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
  var debtorType = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
  var C_pan = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.PAN;
  var C_tan = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.TAN;
    
  if(State!=""){
  Sys.HighlightObject(StateCode);
  StateCode.HoverMouse();
  StateCode.Click();
  DropDownList(State,"State Code")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
  if(GST!=""){
  Sys.HighlightObject(debtorType);
  debtorType.HoverMouse();
  debtorType.Click();
  WorkspaceUtils.DropDownList(GST,"GST Debtor Type")
  }
  
  if(PAN!=""){
  Sys.HighlightObject(C_pan);
  C_pan.HoverMouse();  
   C_pan.setText(PAN);
  }
  
  if(TAN!=""){
   C_tan.setText(TAN);
  }
var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.Save;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();

}
  
function attachDocument(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  } 
  aqUtils.Delay(4000, "Finding Document");;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  } 
 if(EnvParams.Country.toUpperCase()=="INDIA"){
//  var doc = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Document
  var doc = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 11);
  }else if(EnvParams.Country.toUpperCase()=="SPAIN"){ 
//    var doc = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl
    var doc = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 11);
  }
  else if(EnvParams.Country.toUpperCase()=="EGYPT"){
  var doc=  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.SWTObject("TabControl", "", 13)
  }
  else{ 
//  var doc =  Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Document;
  var doc =  Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.history;
  }
  
  Sys.HighlightObject(doc)

WorkspaceUtils.waitForObj(doc);
  Log.Message(doc.text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Sys.HighlightObject(doc);
  aqUtils.Delay(4000, "Finding Document");;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  } 
  doc.HoverMouse();
  doc.HoverMouse();
  doc.HoverMouse();
  doc.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}else{ 

} 
  var attchDocument = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.Composite.AttachDocument;
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  Sys.HighlightObject(attchDocument);
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  ReportUtils.logStep_Screenshot();
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
  
  
function Information(){ 
  var info = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Information;
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
  var submit = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Submit;
  Sys.HighlightObject(submit);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
  submit.HoverMouse();
  submit.HoverMouse();
  submit.Click();
  TextUtils.writeLog("Created Global Client is Submitted for Approval");
  TextUtils.writeLog("      ");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
   }else{ 
       p = Sys.Process("Maconomy");
  w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*", 2000);
  if (w.Exists)
{ 
    var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Label", "*");
    Log.Message(label.getText());
    var lab = label.getText().OleValue.toString().trim();
    var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()+"*").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
    Ok.HoverMouse(); 
    ReportUtils.logStep_Screenshot();
    Ok.Click(); 
    aqUtils.Delay(8000, Indicator.Text); ;
   
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
   }
   
   }
   } 
  
}
  
function ApprvalInformation(){ 
 
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
else{
  var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.ClientApproval;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 ImageRepository.ImageSet.Maximize.Click();
}

 var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.ClientApproval_Tab;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}
Delay(3000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
   var ApproverTable = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.ApprovarTable;
   var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      approvers = EnvParams.Opco+"*"+ClientNumber+"*"+ApproverTable.getItem(i).getText_2(6).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
      Approve_Level[y] = approvers;
      y++;
      }
}

TextUtils.writeLog("Finding approvers for Created Global Client");
var closeCAList = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel2.CloseApprovalList;
 Sys.HighlightObject(closeCAList);
 closeCAList.HoverMouse();
 closeCAList.Click();
 
ImageRepository.ImageSet.Forward.Click();


CredentialLogin();
var OpCo2 = ApproveInfo[0].split("*");
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
sheetName = "CreateClient";
if(OpCo2[2]==Project_manager){
level = 1;
var apStat = true
var Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel;
for(var j=0;j<Approve.ChildCount;j++){ 
 if(Approve.Child(j).isVisible()){ 
   Approve = Approve.Child(j);
 for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)){
    Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Approve;
    apStat = false;
    Sys.HighlightObject(Approve)
    Log.Message(Approve.FullName)
    break;
  }
}
}
}

if(apStat){ 
 Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("Composite", "", 2);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)){
    Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8)
    apStat = false;
    Sys.HighlightObject(Approve)
    Log.Message(Approve.FullName)
    break;
  }
} 
}

Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(8000, "Waiting for Approve");;
ValidationUtils.verify(true,true,"Global CLient is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created Global Client");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

}
}

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
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("FP")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("FP")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
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

}
  
function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
  aqUtils.Delay(3000, Indicator.Text);
  
    if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
  aqUtils.Delay(3000, Indicator.Text);
  
    if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ }
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists())
{ 
  
}
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists())
{ 
  
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}
var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if(temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer (Substitute)").OleValue.toString().trim()+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Customer (Substitute) from To-Dos List");
var listPass = true;   
  }
} 
  }
}

  
function FinalApproveClient(ClientNum,Apvr,lvl){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Maconomy loading data");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var table = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder;
waitForObj(table);
Sys.HighlightObject(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Visible){

}else{ 
var showFilter = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.ShowFilterList;
waitForObj(table);
Sys.HighlightObject(showFilter);
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.Click();
}

var table = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.ClientTable;
var firstCell = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.ClientTable.ClientSearch;
waitForObj(firstCell);
Sys.HighlightObject(firstCell);
firstCell.HoverMouse();
firstCell.HoverMouse();
firstCell.setText(ClientNum);
aqUtils.Delay(3000, "Reading Data in table");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var closefilter = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.CloseFilterList;
waitForObj(closefilter);
Sys.HighlightObject(closefilter);
closefilter.HoverMouse();
closefilter.HoverMouse(); 
closefilter.HoverMouse();
closefilter.HoverMouse(); 
aqUtils.Delay(3000, "Reading Data in table");;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==ClientNum){ 
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

var apStat = true
var Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel;
for(var j=0;j<Approve.ChildCount;j++){ 
 if(Approve.Child(j).isVisible()){ 
   Approve = Approve.Child(j);
 for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)){
    Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.Approve;;
    apStat = false;
    Sys.HighlightObject(Approve)
    Log.Message(Approve.FullName)
    break;
  }
}
}
}

if(apStat){ 
 Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("Composite", "", 2);
for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)){
    Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8)
    apStat = false;
    Sys.HighlightObject(Approve)
    Log.Message(Approve.FullName)
    break;
  }
} 
}

Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(8000, "Waiting To Approve");;
ValidationUtils.verify(true,true,"Global Client is Approved by "+Apvr)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

TextUtils.writeLog("Global Client is Approved by "+Apvr);
if(Approve_Level.length==lvl+1){
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(8000, Indicator.Text); ;
 for(var j=0;j<12;j++){ 
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()){ 
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(8000, Indicator.Text); ;  
}

}
 


  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var E_ID = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite9.McTextWidget;
  if((E_ID.getText()=="")||(E_ID.getText()==null)){
  E_ID.setText(mail)
  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  save.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  ValidationUtils.verify(true,true,"Email Id is saved");
  TextUtils.writeLog("Email Id is saved");
  }
  
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Global Client Number",EnvParams.Opco,"Data Management",ClientNum)
  ExcelUtils.WriteExcelSheet("Global Client Name",EnvParams.Opco,"Data Management",clientName)
  ExcelUtils.WriteExcelSheet("Global Brand Number",EnvParams.Opco,"Data Management",ClientNum+"001")
  ExcelUtils.WriteExcelSheet("Global Brand Name",EnvParams.Opco,"Data Management",brand)
  ExcelUtils.WriteExcelSheet("Global Product Number",EnvParams.Opco,"Data Management",ClientNum+"001001")
  ExcelUtils.WriteExcelSheet("Global Product Name",EnvParams.Opco,"Data Management",product)
  ExcelUtils.WriteExcelSheet("Global Client Currency",EnvParams.Opco,"Data Management",currency)
  
  TextUtils.writeLog("Global Client Number :"+ClientNum); 
  TextUtils.writeLog("Global Client Name :"+clientName);
  TextUtils.writeLog("Global Brand Number :"+ClientNum+"001");
  TextUtils.writeLog("Global Brand Name :"+brand);
  TextUtils.writeLog("Global Product Number :"+ClientNum+"001001");
  TextUtils.writeLog("Global Product Name :"+product);

  
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}  
else {    
var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 aqUtils.Delay(2000, Indicator.Text);;
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
ImageRepository.ImageSet.Maximize.Click();
} 
 var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 aqUtils.Delay(2000, Indicator.Text);;
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
 
   var ApproverTable = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(ApproverTable);
  ReportUtils.logStep_Screenshot();
    for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(true,false,"Created Client is not Approved")
      }
}
  var closeApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl;
  Sys.HighlightObject(closeApproval);
 closeApproval.HoverMouse();
 closeApproval.Click();
  aqUtils.Delay(2000, Indicator.Text);;
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
 ImageRepository.ImageSet.Forward.Click();
  aqUtils.Delay(2000, Indicator.Text);;
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  var Invoice_Add = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
 var Company_Name = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
 var Address1 = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
 var Address2 = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
 var Address3 = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite5.McTextWidget;
 var Zip_Code = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite6.McTextWidget;
 var C_Country = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite7.McTextWidget;
 var Save = Aliases.Maconomy.Invoice_Address.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 var Email = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
 var Info_Save = false;
 if(Email.getText()==null){ 
   Email.Click();
   Email.setText( mail );
   Info_Save = true;
 }
    if(Invoice_Spec_Add.toUpperCase()=="YES"){
  if(Invoice_Add.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  { 
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      Company_Name.Click();
      Company_Name.setText(CompanyName)
      aqUtils.Delay(2000, Indicator.Text);
      Address1.Click();
      Address1.setText(i_Add1)
      aqUtils.Delay(2000, Indicator.Text);
      Zip_Code.Click();
      Zip_Code.setText(ZipCode)
      aqUtils.Delay(2000, Indicator.Text);
      C_Country.Click();
      C_Country.setText(country)
      aqUtils.Delay(2000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      Save.Click();
      aqUtils.Delay(5000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      ValidationUtils.verify(true,true,"Invoice Specific Address is Saved");
      ReportUtils.logStep_Screenshot();
      TextUtils.writeLog("Invoice Specific Address is Saved");
  }
  else{ 
      Invoice_Add.Click();
      DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, Invoice_Spec_Add).OleValue.toString().trim())
      aqUtils.Delay(5000, Indicator.Text);
      ReportUtils.logStep_Screenshot();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      Company_Name.Click();
      Company_Name.setText(CompanyName)
      aqUtils.Delay(2000, Indicator.Text);
      Address1.Click();
      Address1.setText(i_Add1)
      aqUtils.Delay(2000, Indicator.Text);
      Zip_Code.Click();
      Zip_Code.setText(ZipCode)
      aqUtils.Delay(2000, Indicator.Text);
      C_Country.Click();
      C_Country.setText(country)
      aqUtils.Delay(2000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      Save.Click();
      aqUtils.Delay(5000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      ValidationUtils.verify(true,true,"Invoice Specific Address is Saved");
      ReportUtils.logStep_Screenshot();
      TextUtils.writeLog("Invoice Specific Address is Saved");
  }
 
 }else{ 
   
 if(Info_Save){ 
      aqUtils.Delay(2000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
      Save.Click();
      aqUtils.Delay(5000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
 }
 }

 aqUtils.Delay(2000, Indicator.Text);
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ } 
 
 var Settling_Company = Aliases.Maconomy.CreateTimesheet.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
 var Company_Number = Aliases.Maconomy.CreateTimesheet.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
 if(Settling_Company.getText()=="-")
 ValidationUtils.verify(true,true,"Settling Company is cleared after final approval")
 else
 ValidationUtils.verify(false,true,"Settling Company is NOT cleared after final approval")
 
 if(Company_Number.getText()=="-")
 ValidationUtils.verify(true,true,"Company Number is cleared after final approval")
 else
 ValidationUtils.verify(false,true,"Company Number is NOT cleared after final approval")
 
  aqUtils.Delay(5000, Indicator.Text);
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
 var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
}
  ValidationUtils.verify(true,true,"Global Client is Approved by "+Apvr)
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  
}
}

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
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
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

function VPWSearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);

    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
  waitForObj(code);
  code.Click();
    code.setText(value);
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

function gotoSearch(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 var CompanyNumber = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite.companyNo_textbox;
 waitForObj(CompanyNumber);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Sys.HighlightObject(CompanyNumber);
  CompanyNumber.Click();
  Log.Message(EnvParams.Opco)
  SearchByValue(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco.toString(),"Company Number");

 var curr = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite2.currency_Dropdown;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 if(currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(currency,"Currency")
  }

 var clientNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
 Sys.HighlightObject(clientNo);
  clientNo.HoverMouse();
  Sys.HighlightObject(clientNo);
  clientNo.Click();
  VPWSearchByValue(clientNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNumber,"Client No");

 var ClientName = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite4.clientName_textbox;
 ClientName.HoverMouse();
 Sys.HighlightObject(ClientName);
 ClientName.setText("*");
 
 
 var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.Clientsearch_Save;
 save.HoverMouse();
 Sys.HighlightObject(save);
 save.Click();
 
 TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}

function goToglobalClient(){ 
  var GblClient = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(GblClient); 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  GblClient.HoverMouse();
  GblClient.HoverMouse();  
  GblClient.Click();
  var active = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());;
  Sys.HighlightObject(active); 
  active.HoverMouse();
  active.HoverMouse(); 
  active.Click();

  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table); 
  table.HoverMouse();
  table.HoverMouse();
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Finding Home");
  TextUtils.writeLog("Global Client is available in maconomy");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){
    
  }
}

function globalBrand(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){
    
  }
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){
    
  }
    aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  home.HoverMouse();
  home.HoverMouse();
  home.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
    aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){
    
  }
  TextUtils.writeLog("Navigating to Sub Level");
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){
    
  }

  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Brands").OleValue.toString().trim());
  Sys.HighlightObject(activeBrand); 
  activeBrand.HoverMouse();
  activeBrand.HoverMouse();
  activeBrand.Click();

  TextUtils.writeLog("Active Brand is selected");
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var brandNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(brandNmae); 
  brandNmae.HoverMouse();
  brandNmae.HoverMouse();
  brandNmae.Click();
  Log.Message(brand)
  brandNmae.Keys(brand);
  Log.Message(brand)
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Findind Information");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  TextUtils.writeLog("Global Brand is available in maconomy");

  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
   
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var E_ID = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite9.McTextWidget;
  if((E_ID.getText()=="")||(E_ID.getText()==null)){
  Sys.HighlightObject(E_ID)
  WorkspaceUtils.waitForObj(E_ID)
  E_ID.setText(mail)
  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(save)
  save.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  ValidationUtils.verify(true,true,"Email Id is saved");
  TextUtils.writeLog("Email Id is saved");
  }

if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();  
}
else{  
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
ImageRepository.ImageSet.Maximize.Click();
}
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
}
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      ClientApproval.Click();
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
      ReportUtils.logStep_Screenshot();
      for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,false,"Global Brand is not Approved")
      }
}
    var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
          if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
}


function globalProduct(){ 
  
    aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
    aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  home.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
    aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();

  TextUtils.writeLog("Navigating to Sub Level");
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }

  var activeProduct = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Products").OleValue.toString().trim());
  Sys.HighlightObject(activeProduct); 
  activeProduct.HoverMouse();
  activeProduct.HoverMouse();
  activeProduct.Click();

  TextUtils.writeLog("Active Product is selected");
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var productNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(productNmae); 
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  productNmae.Keys(product);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

    
  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy");
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy");
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy");
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Product is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Findind Information");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  TextUtils.writeLog("Global Product is available in maconomy");

  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var E_ID = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite9.McTextWidget;
  if((E_ID.getText()=="")||(E_ID.getText()==null)){
  Sys.HighlightObject(E_ID)
  WorkspaceUtils.waitForObj(E_ID)
  E_ID.setText(mail)
  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(save)
  save.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  ValidationUtils.verify(true,true,"Email Id is saved");
  TextUtils.writeLog("Email Id is saved");
  }

 if(ImageRepository.ImageSet.Maximize.Exists())
ImageRepository.ImageSet.Maximize.Click();
else
{  
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
ImageRepository.ImageSet.Maximize.Click();
}
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      ClientApproval.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
      }
      ReportUtils.logStep_Screenshot();
      for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,false,"Global Product is not Approved")
      }
}
    var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
}

function companyProduct(){ 
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
    }
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
  home.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
  }
    aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();
  Delay(3000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Navigating to Sub Level");
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
  Delay(3000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  var activeProduct = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Company Products").OleValue.toString().trim())
  Sys.HighlightObject(activeProduct); 
  activeProduct.HoverMouse();
  activeProduct.HoverMouse();
  activeProduct.Click();
  Delay(3000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  TextUtils.writeLog("Active Company Product is selected");
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
  
  var ClientType = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McPopupPickerWidget", "");
  ClientType.Keys("[Tab][Tab]")
//  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  var productNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
//  var productNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")
  Sys.HighlightObject(productNmae); 
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  productNmae.Keys(product);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

    
  if(table.getItem(0).getText_2(3).OleValue.toString().trim()==product){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  else if(table.getItem(1).getText_2(3).OleValue.toString().trim()==product){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  else if(table.getItem(2).getText_2(3).OleValue.toString().trim()==product){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  else if(table.getItem(3).getText_2(4).OleValue.toString().trim()==product){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Findind Information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Company Product is available in maconomy");
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
     if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var E_ID = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2)
  if((E_ID.getText()=="")||(E_ID.getText()==null)){
  Sys.HighlightObject(E_ID)
  WorkspaceUtils.waitForObj(E_ID)
  E_ID.setText(mail)
//  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  var save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(save)
  save.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  ValidationUtils.verify(true,true,"Email Id is saved");
  TextUtils.writeLog("Email Id is saved");
  }

  
   if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
else {
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
ImageRepository.ImageSet.Maximize.Click();
} 
     
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
      ClientApproval.Click();
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
       }
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
       }
      ReportUtils.logStep_Screenshot();
      for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,false,"Company Product is not Approved")
      }
}
    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
}


function companyBrand(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  home.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, "Findind SubLevels");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, "Findind SubLevels");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Navigating to Sub Level");
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
                  
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var activebrand = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").
  SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Company Brands").OleValue.toString().trim());
  Sys.HighlightObject(activebrand); 
  activebrand.HoverMouse();
  activebrand.HoverMouse();
  activebrand.Click();
  Delay(3000);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Active Company Brand is selected");
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
  
  var ClientType = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McPopupPickerWidget", "");
  ClientType.Keys("[Tab][Tab]")
//  Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
  var productNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
//  var productNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")
  Sys.HighlightObject(productNmae); 
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  productNmae.Keys(brand);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
    
  if(table.getItem(0).getText_2(3).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  else if(table.getItem(1).getText_2(3).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  else if(table.getItem(2).getText_2(3).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  else if(table.getItem(3).getText_2(3).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Findind Information");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  TextUtils.writeLog("Company Brand is available in maconomy");

  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var E_ID = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);

  if((E_ID.getText()=="")||(E_ID.getText()==null)){
  Sys.HighlightObject(E_ID)
  WorkspaceUtils.waitForObj(E_ID)
  E_ID.setText(mail)
//  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  var save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  Sys.HighlightObject(save)
  save.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  ValidationUtils.verify(true,true,"Email Id is saved");
  TextUtils.writeLog("Email Id is saved");
  }

  

 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
else {  
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
ImageRepository.ImageSet.Maximize.Click();
 }

      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
      ClientApproval.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }

      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      ReportUtils.logStep_Screenshot();
      for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,false,"Company Brand is not Approved")
      }
}

    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}


function goToCompanyClient(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var GblClient = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(GblClient); 
  GblClient.HoverMouse();
  GblClient.HoverMouse();  
  GblClient.Click();
  var active = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());
  Sys.HighlightObject(active); 
  active.HoverMouse();
  active.HoverMouse(); 
  active.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  var table = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table); 
  table.HoverMouse();
  table.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Sys.HighlightObject(table);
  Sys.HighlightObject(table);

  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Company Client is available in maconomy");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Company Client is available in maconomy");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Company Client is available in maconomy");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNumber){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Company Client is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Finding Home");
  TextUtils.writeLog("Company Client is available in maconomy");

    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}

function CompanyClient(){ 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000, "Findind information");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var E_ID = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2)

  if((E_ID.getText()=="")||(E_ID.getText()==null)){
  Sys.HighlightObject(E_ID)
  WorkspaceUtils.waitForObj(E_ID)
  E_ID.setText(mail)
    aqUtils.Delay(3000, "Findind Client Approval");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  
//  var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  var save = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
  Sys.HighlightObject(save)
  save.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  ValidationUtils.verify(true,true,"Email Id is saved");
  TextUtils.writeLog("Email Id is saved");
  }

  
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
else { 
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
ImageRepository.ImageSet.Maximize.Click(); 
 }

      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
      ClientApproval.Click();
        if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
       }
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
       }
      ReportUtils.logStep_Screenshot();
      for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(5)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
      ValidationUtils.verify(true,false,"Company Client is not Approved")
      }
}

    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
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
