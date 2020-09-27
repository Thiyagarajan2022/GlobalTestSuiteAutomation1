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
var sheetName = "CreateGlobalBrand";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var clientName,brandname,defaultname,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product,State,GST,PAN,TAN,SII_Tax,TIN ="";
var ClientNumber = "";
var client1="";
var Language = "";
var BrdNum = "";
function CreateGlobalBrand(){
TextUtils.writeLog("Global Brand Creation Started"); 
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateGlobalBrand";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
clientName,brandname,defaultname,strt1,strt2,P_code,P_District,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,Fax,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product,State,GST,PAN,TAN,SII_Tax,TIN ="";
ClientNumber = "";
Approve_Level = [];
BrdNum = "";
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Client Creation started::"+STIME);
TextUtils.writeLog("Execution Started :"+STIME);
getDetails();

//clientName = "1707_AutoClient 25February2020 10:41:17";
//brand = "AutoGlobalBrand 25February2020 10:41:17";
//product = "AutoGlobalProduct 25February2020 10:41:17" ;
//ClientNumber = "107743"

gotoMenu(); 
gotoClientSearch();
globalClient(); 
newGlobalClient();
GlobalClient_Screen2();
test();
brand1();
if(EnvParams.Country.toUpperCase()=="INDIA"){
Runner.CallMethod("IND_CreateGlobalBrand.indiaSpecific",State,GST,PAN,TAN,TIN);
}
if(EnvParams.Country.toUpperCase()=="SPAIN"){
Runner.CallMethod("SPA_GlobalBrand.spainSpecific",SII_Tax);
}
//attachDocument();
Information();
ApprvalInformation();
//ClientDueDiligencePolicy();
//globalClientTable();

//attachDocument();
//Information();
//ApprvalInformation();
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
}
WorkspaceUtils.closeAllWorkspaces();
//gotoMenu(); 
//gotoSearch();
//goToglobalClient();
//globalBrand();
//WorkspaceUtils.closeAllWorkspaces();
//gotoMenu(); 
//gotoSearch();
//goToglobalClient();
//globalProduct();
//WorkspaceUtils.closeAllWorkspaces();
//gotoMenu(); 
//gotoSearch();
//goToCompanyClient();
//CompanyClient();
//WorkspaceUtils.closeAllWorkspaces();
//gotoMenu(); 
//gotoSearch();
//goToglobalClient();
//companyBrand();
//WorkspaceUtils.closeAllWorkspaces();
//gotoMenu(); 
//gotoSearch();
//goToglobalClient();
//companyProduct();
//WorkspaceUtils.closeAllWorkspaces();


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

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receivable Menu");
TextUtils.writeLog("Entering into Client Management from Accounts Receivable Menu");
}


function getDetails(){ 
Indicator.PushText("Reading Data from Excel");
ExcelUtils.setExcelName(workBook, sheetName, true);

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  ClientNo = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
  if((ClientNo=="")||(ClientNo==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  }
  
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  clientName = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
  if((clientName=="")||(clientName==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  clientName = ExcelUtils.getRowDatas("Client Name",EnvParams.Opco)
  }
  if((clientName==null)||(clientName=="")){ 
  ValidationUtils.verify(false,true,"Client Name is Needed to Create a Client");
  }
Log.Message(clientName)

ExcelUtils.setExcelName(workBook, sheetName, true);
brandname = ExcelUtils.getRowDatas("Brand Name",EnvParams.Opco)

if((brandname==null)||(brandname=="")){ 
ValidationUtils.verify(false,true,"Brand Name is Needed to Create a Client");
}


Log.Message(brandname)
ExcelUtils.setExcelName(workBook, sheetName, true);
defaultname = ExcelUtils.getRowDatas("Product Name",EnvParams.Opco)
if((defaultname==null)||(defaultname=="")){ 
ValidationUtils.verify(false,true,"Product Name is Needed to Create a Client");
}


Log.Message(defaultname)

clientlan = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((clientlan==null)||(clientlan=="")){ 
ValidationUtils.verify(false,true,"Language is Needed to Create a Client");
}
Log.Message(clientlan)

currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((currency==null)||(currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Create a Client");
}
Log.Message(currency)


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
ValidationUtils.verify(false,true,"SII Tax Group is Needed to Create a Global Brand");
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
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 var CompanyNumber = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
 waitForObj(CompanyNumber);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  Sys.HighlightObject(CompanyNumber);
  CompanyNumber.Click();
  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 var curr = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 if(currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(currency,"Currency")
  }
//  aqUtils.Delay(2000, Indicator.Text);
var ClientNumber = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget;
  if(ClientNo!=""){
  ClientNumber.Click();
  WorkspaceUtils.VPWSearchByValue(ClientNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client").OleValue.toString().trim(),ClientNo,"Client Number");
  }

 var ClientName = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
 ClientName.HoverMouse();
 Sys.HighlightObject(ClientName);
 ClientName.setText("*");
 //ClientName.setText(clientName+" "+STIME);
 
 
 var save = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
 save.HoverMouse();
 Sys.HighlightObject(save);
 save.Click();
// aqUtils.Delay(5000, Indicator.Text);
 
 TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
   if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
}


function globalClient(){ 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var GblClient = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  GblClient.HoverMouse();
  Sys.HighlightObject(GblClient);
  GblClient.Click();
//  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var AllClients = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());
  AllClients.Click();
  AllClients.HoverMouse();
  AllClients.HoverMouse();
  AllClients.HoverMouse();
  
  aqUtils.Delay(3000, "Reading from Global Client table");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNo){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
  }
  
  aqUtils.Delay(5000, "Playback");
  TextUtils.writeLog("Global Client is available in maconomy to Amend");
  
  
  }



function newGlobalClient(){ 
  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
var home= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
// var ClientName = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.ClientName_textbox;
 waitForObj(home);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 home.Click();
 Sys.HighlightObject(home);
// ClientName.setText(clientName.toString().trim()+" "+STIME);
// clientName = clientName.toString().trim()+" "+STIME;
 
var sublevel= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
// var ClientName = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.ClientName_textbox;
 waitForObj(sublevel);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 sublevel.Click();
 Sys.HighlightObject(sublevel);
 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 var glbclient= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
  waitForObj(glbclient);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
 glbclient.Click();
 Sys.HighlightObject(glbclient);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var newglbbrand= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  waitForObj(newglbbrand);
  newglbbrand.Click();
 Sys.HighlightObject(newglbbrand);
 aqUtils.Delay(5000, "Playback");
var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
//   var cancel = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
 waitForObj(cancel)
 
var brandname1=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
 // waitForObj(brandname1);
  brandname1.Click();
Sys.HighlightObject(brandname1);
//brandname1.setText(brandname);

 brandname1.setText(brandname+" "+STIME);
 brandname = brandname+" "+STIME;

var defaultname1=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
 // waitForObj(brandname1);
  defaultname1.Click();
Sys.HighlightObject(defaultname1);
defaultname1.setText(defaultname);

//Default Name
 
var next=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
 next.HoverMouse();
 Sys.HighlightObject(next);
 ReportUtils.logStep_Screenshot() ;
 next.Click();
 
//Global Client - Client Information Card
}

function GlobalClient_Screen2(){ 
  var CompanyNumber = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McValuePickerWidget;
 waitForObj(CompanyNumber);
 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
 waitForObj(cancel)
  CompanyNumber.Click();
    Sys.HighlightObject(CompanyNumber);
  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");


  var C_Language = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McPopupPickerWidget;
  if(clientlan!=""){
  C_Language.Click();
  WorkspaceUtils.DropDownList(clientlan,"Language")
  }
//  aqUtils.Delay(2000, Indicator.Text);
  
  var Attn = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite3.McValuePickerWidget;
  Attn.HoverMouse();
  Sys.HighlightObject(Attn);
  Attn.setText(attn);  
  
  var C_Email  = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite4.McTextWidget;
//  var Eml_split1 = mail.OleValue.toString().trim().substring(0,mail.OleValue.toString().trim().indexOf("@"));
//  var Eml_split2 = mail.OleValue.toString().trim().substring(mail.OleValue.toString().trim().indexOf("@"));
//  Eml_split1 = Eml_split1 +" "+STIME;
//  Eml_split1 = Eml_split1.replace(/[_: ]/g,"");
//  mail = Eml_split1+Eml_split2
  C_Email.setText(mail);
  
  var C_phone = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite5.McTextWidget;
  C_phone.setText(phone); 
  
  var C_AcctDir = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite6.McValuePickerWidget;
  if(AccDir!=""){
  C_AcctDir.HoverMouse();
  Sys.HighlightObject(C_AcctDir);
  C_AcctDir.Click();
  WorkspaceUtils.SearchByValue(C_AcctDir,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee").OleValue.toString().trim(),AccDir,"Acct Director No");
  }
  
  var C_PaymentTerm = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite7.McPopupPickerWidget;
  if(payterm!=""){
  Sys.HighlightObject(C_PaymentTerm);
  C_PaymentTerm.Click();
  WorkspaceUtils.DropDownList(payterm,"Payment Terms")
  }
//  aqUtils.Delay(2000, Indicator.Text);
  
  var C_companyTaxCode = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite8.McPopupPickerWidget;
  if(Comtaxcode!=""){
  C_companyTaxCode.HoverMouse();
  Sys.HighlightObject(C_companyTaxCode);
  C_companyTaxCode.Click();
  WorkspaceUtils.DropDownList(Comtaxcode,"Company Tax Code");
  }
//  aqUtils.Delay(2000, Indicator.Text);
  
  var C_JobPriceList = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite9.McValuePickerWidget;
  if(sales!=""){
  Sys.HighlightObject(C_JobPriceList);
  C_JobPriceList.Click();
  WorkspaceUtils.SearchByValue(C_JobPriceList,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Price List").OleValue.toString().trim(),sales,"Job Price List Sales");
         }  
         
//    aqUtils.Delay(2000, Indicator.Text);
  
 var Next = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Create").OleValue.toString().trim());
 Sys.HighlightObject(Next);
 Next.HoverMouse();
 ReportUtils.logStep_Screenshot() ;
 Next.Click();
    aqUtils.Delay(7000, "Finding pop-up");

}

//
function test()
{
  
// if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Client Information Card").OleValue.toString().trim())    
//    {
      
//p = Sys.Process("Maconomy");
//w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Client Information Card").OleValue.toString().trim(), 2000);
//if (w.Exists)
//{
var w = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Client Information Card").OleValue.toString().trim());
var Label = w.SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
//}
aqUtils.Delay(7000, "Finding pop-up");
p = Sys.Process("Maconomy");
w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Client Information Card").OleValue.toString().trim(), 2000);
if (w.Exists)
{
var Label = w.SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();
}

//    var button = Aliases.Maconomy.CreateGlobalBrand3.Composite.Button;
//    //Aliases.CompanyRegistrationAlreadyUsED
//      var label =Aliases.Maconomy.CreateGlobalBrand3.SWTObject("Label", "*").WndCaption;
//     //  NameMapping.Sys.Maconomy.Shell.SWTObject("Label", "*").WndCaption;
//      Log.Message(label );
//       button.HoverMouse();
//    // ReportUtils.logStep_Screenshot("");
//      button.Click();
      aqUtils.Delay(5000, "Finding pop-up");
//  }
     
    
 if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Global Client - Client Information Card")    
    { 
  var button1=Aliases.Maconomy.CreateGlobalBrand3.Composite.Button;
  var label =Aliases.Maconomy.CreateGlobalBrand3.SWTObject("Label", "*").WndCaption;
  Log.Message(label );
       button1.HoverMouse();
    // ReportUtils.logStep_Screenshot("");
      button1.Click();
      aqUtils.Delay(5000, "Finding pop-up");
}
}



function brand1(){
//var name1=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
//name1.Click();
//  name1.Keys(brandname); 

var blocked = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blocked").OleValue.toString().trim())
blocked.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  aqUtils.Delay(3000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
var name1=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
name1.Click();
  name1.Keys(brandname); 
var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 Sys.HighlightObject(table);
//Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
 
  aqUtils.Delay(4000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==brandname){
  BrdNum = table.getItem(0).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==brandname){
  BrdNum = table.getItem(1).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==brandname){
  BrdNum = table.getItem(2).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==brandname){
  BrdNum = table.getItem(3).getText_2(0).OleValue.toString().trim();
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Global Brand is available in maconomy");
  }
  
  aqUtils.Delay(5000, Indicator.Text);

  TextUtils.writeLog("Global Brand is available in maconomy");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  aqUtils.Delay(5000, Indicator.Text);
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
  DropDownList(State.OleValue.toString().trim(),"State Code")
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
   C_pan.setText(PAN.OleValue.toString().trim());
  }
  
  if(TAN!=""){
   C_tan.setText(TAN.OleValue.toString().trim());
  }
var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.Save;
Sys.HighlightObject(save);
save.HoverMouse();
save.Click();

}
  
function attachDocument(){ 
//  var doc = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Document;
//  var doc = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Document
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 if(EnvParams.Country.toUpperCase()=="INDIA"){
  var doc = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Document
  }
  else{ 
  var doc =  Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Document
  }
  Sys.HighlightObject(doc);
  doc.HoverMouse();
  doc.HoverMouse();
  doc.HoverMouse();
  doc.Click();
  var attchDocument = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.Composite.AttachDocument;
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  Sys.HighlightObject(attchDocument);
  attchDocument.HoverMouse();
  attchDocument.HoverMouse();
  ReportUtils.logStep_Screenshot();
  attchDocument.Click();
  aqUtils.Delay(4000, Indicator.Text);;
  var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
  dicratory.Keys(workBook);
  var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
  Sys.HighlightObject(opendoc);
  opendoc.HoverMouse();
  ReportUtils.logStep_Screenshot();
  opendoc.Click();
  aqUtils.Delay(2000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}
  
  
function Information(){ 
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN")){
  var info = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl
  }else{ 
  var info = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl2;
}
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
if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN")){
  var submit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2.SingleToolItemControl2;
  }else{
  var submit = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
}
  Sys.HighlightObject(submit);
  submit.HoverMouse();
  submit.HoverMouse();
  submit.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
}
  
function ApprvalInformation(){ 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN")){
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;;
 }else{
 var ClientApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl
}
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

// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl
 var ClientApproval = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}


   var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//   var ApproverTable = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
   var y=0;
  for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";
      if(ApproverTable.getItem(i).getText_2(3)!="Approved"){
      approvers = EnvParams.Opco+"*"+BrdNum+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(5).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
//      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
      Approve_Level[y] = approvers;
      y++;
      }
}

TextUtils.writeLog("Finding approvers for Created Global Client");
var closeCAList = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl;
//var closeCAList = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
 Sys.HighlightObject(closeCAList);
 closeCAList.HoverMouse();
 closeCAList.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
ImageRepository.ImageSet.Forward.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

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
Project_manager = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
Project_manager = Project_manager.substring(Project_manager.indexOf(" - ")+3);
sheetName = "CreateGlobalBrand";
if(OpCo2[2]==Project_manager){
level = 1;


//if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN")){
//var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2
//}else{
////var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl
//var Approve = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite;
//}
//
// for(var i=0;i<Approve.ChildCount;i++){ 
//  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
//    Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.Approve;;
//    Sys.HighlightObject(Approve)
//    Log.Message(Approve.FullName)
//    break;
//  }
//}



    var ChildCount = 0;
    var Add = [];
   var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
                
      for(var ip=0;ip<Parent.ChildCount;ip++){ 
     var PChild = Parent.Child(ip);
        if((PChild.isVisible()) && (PChild.ChildCount==1)){
        Add[ChildCount] = PChild;
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
Sys.HighlightObject(Approve);
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(8000, "Waiting for Approve");;
 for(var j=0;j<12;j++){ 
    var p = Sys.Process("Maconomy");
    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()){ 
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;  
}
}
ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created Budget");
//aqUtils.Delay(8000, Indicator.Text);;
}
}
//var Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Approve;
// Sys.HighlightObject(Approve);
// Approve.HoverMouse();
// Approve.Click();
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

  
function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
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
  aqUtils.Delay(1000, Indicator.Text);

var refresh= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
  
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
//var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
//}
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
//var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
//var Client_Managt=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;

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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){
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
//aqUtils.Delay(5000, Indicator.Text);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//aqUtils.Delay(2000, Indicator.Text);
//ImageRepository.ImageSet.Show_Filter.Click();
//}

var GBName = "";
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
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
var closefilter = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite2.CloseFilterList;
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
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==ClientNum){ 
    GBName = table.getItem(v).getText_2(2).OleValue.toString().trim();
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
//NameMapping.Sys.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3
var Approve = NameMapping.Sys.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel;
for(var j=0;j<Approve.ChildCount;j++){ 
 if(Approve.Child(j).isVisible()){ 
   Approve = Approve.Child(j);
 for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
//    Approve = NameMapping.Sys.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
    Approve = Approve.Child(i);
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
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
//    Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 8)
    Approve = Approve.Child(i);
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
aqUtils.Delay(8000, Indicator.Text);;
TextUtils.writeLog("Global Client is Approved by "+Apvr);
if(Approve_Level.length==lvl+1){
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
aqUtils.Delay(4000, Indicator.Text); ;
Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
aqUtils.Delay(4000, Indicator.Text); ;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 for(var j=0;j<12;j++){ 
    var p = Sys.Process("Maconomy");
    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()){ 
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;  
}
 
    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type (Substitute)").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type (Substitute)").OleValue.toString().trim()){ 
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type (Substitute)").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type (Substitute)").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;  
}

    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer").OleValue.toString().trim()){ 
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;  
}

    var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer (Substitute)").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer (Substitute)").OleValue.toString().trim()){ 
var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer (Substitute)").OleValue.toString().trim()).SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer (Substitute)").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse(); 
ReportUtils.logStep_Screenshot();
Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;  
}


 }
 
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Global Brand No",EnvParams.Opco,"Data Management",ClientNum)
  ExcelUtils.WriteExcelSheet("GlobalBrand Name",EnvParams.Opco,"Data Management",GBName)
  TextUtils.writeLog("Global Brand No :"+ClientNum); 
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
// }
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.TabControl;
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}

//   var ApproverTable = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
var ApproverTable = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
  Sys.HighlightObject(ApproverTable);
  ReportUtils.logStep_Screenshot();
for(var i=0;i<ApproverTable.getItemCount();i++){   
var approvers="";
if(ApproverTable.getItem(i).getText_2(6)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){
ValidationUtils.verify(true,false,"Created Global Client is not Approved")
}
}
//  var closeApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
var closeApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl;
  Sys.HighlightObject(closeApproval);
 closeApproval.HoverMouse();
 closeApproval.Click();
 if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
 ImageRepository.ImageSet.Forward.Click();
 var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
}
  ValidationUtils.verify(true,true,"Global Brand is Approved by "+Apvr)
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

function gotoSearch(){ 
 var CompanyNumber = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite.companyNo_textbox;
 waitForObj(CompanyNumber);
// var Add_Visible0 = true;
//  while(Add_Visible0){
//    if(CompanyNumber.isEnabled()){
//      CompanyNumber.HoverMouse();
//      Sys.HighlightObject(CompanyNumber);
//      Add_Visible0 = false;
//      }
//  }
  
  Sys.HighlightObject(CompanyNumber);
  CompanyNumber.Click();
  Log.Message(EnvParams.Opco)
  SearchByValue(CompanyNumber,"Company",EnvParams.Opco.toString(),"Company Number");

 var curr = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite2.currency_Dropdown;
 curr.Keys(" ");
 curr.HoverMouse();
 Sys.HighlightObject(curr);
 if(currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(currency.OleValue.toString().trim(),"Currency")
  }
//  aqUtils.Delay(2000, Indicator.Text);
 var clientNo = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
 Sys.HighlightObject(clientNo);
  clientNo.HoverMouse();
  Sys.HighlightObject(clientNo);
  clientNo.Click();
  VPWSearchByValue(clientNo,"Client",ClientNumber,"Client No");

 var ClientName = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.searchClient.Composite.McPaneGui_10.Composite.Composite.searchCriteria.Composite4.clientName_textbox;
 ClientName.HoverMouse();
 Sys.HighlightObject(ClientName);
 ClientName.setText("*");
 
 
 var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.Clientsearch_Save;
 save.HoverMouse();
 Sys.HighlightObject(save);
 save.Click();
// aqUtils.Delay(5000, Indicator.Text);
 
 TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
}

function goToglobalClient(){ 
  var GblClient = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(GblClient); 
  GblClient.HoverMouse();
  GblClient.HoverMouse();  
  GblClient.Click();
  var active = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  Sys.HighlightObject(active); 
  active.HoverMouse();
  active.HoverMouse(); 
  active.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table); 
  table.HoverMouse();
  table.HoverMouse();
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
}

function globalBrand(){ 
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  home.Click();
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  TextUtils.writeLog("Navigating to Sub Level");
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button2;
  Sys.HighlightObject(activeBrand); 
  activeBrand.HoverMouse();
  activeBrand.HoverMouse();
  activeBrand.Click();
//  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Active Brand is selected");
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var brandNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(brandNmae); 
  brandNmae.HoverMouse();
  brandNmae.HoverMouse();
  brandNmae.Click();
  brandNmae.Keys(brand);
  Sys.HighlightObject(table);
//  aqUtils.Delay(4000, Indicator.Text);
    
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

  TextUtils.writeLog("Global Brand is available in maconomy");
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
  
  var ClientNumber1 = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
client1=   ClientNumber1.getText();
   
Log.Message(client1);


// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
// }
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}
//      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl3;
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      ClientApproval.Click();
//      var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      ReportUtils.logStep_Screenshot();
//    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
    var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
}


function globalProduct(){ 
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  home.Click();
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  TextUtils.writeLog("Navigating to Sub Level");
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  //var activeProduct = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  var activeProduct = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button3;
  Sys.HighlightObject(activeProduct); 
  activeProduct.HoverMouse();
  activeProduct.HoverMouse();
  activeProduct.Click();
//  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Active Product is selected");
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var productNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  Sys.HighlightObject(productNmae); 
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  productNmae.Keys(product);
  Sys.HighlightObject(table);
//  aqUtils.Delay(4000, Indicator.Text);
    
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

  TextUtils.writeLog("Global Product is available in maconomy");
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
   


// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
// }
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}

//      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl3;
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      ClientApproval.Click();
//      var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      ReportUtils.logStep_Screenshot();
//    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
    var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
}

function companyProduct(){ 
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  home.Click();
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  TextUtils.writeLog("Navigating to Sub Level");
//  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  //var activeProduct = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  var activeProduct = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - IND CT Clients (TSTAUTO)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Company Products")
  Sys.HighlightObject(activeProduct); 
  activeProduct.HoverMouse();
  activeProduct.HoverMouse();
  activeProduct.Click();
//  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Active Company Product is selected");
  //var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
//  var productNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  var productNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")
  Sys.HighlightObject(productNmae); 
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  productNmae.Keys(product);
  Sys.HighlightObject(table);
//  aqUtils.Delay(4000, Indicator.Text);
    
  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==product){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Company Product is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Findind Information");

  TextUtils.writeLog("Company Product is available in maconomy");
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
   


// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
// var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl

 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
// }
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}

//      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl3;
//      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
      ClientApproval.Click();
      //var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      ReportUtils.logStep_Screenshot();
    //var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
//    var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
}


function companyBrand(){ 
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  Sys.HighlightObject(home); 
  home.HoverMouse();
  home.HoverMouse();
  home.Click();
  var sublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl2;
  Sys.HighlightObject(sublevels); 
  sublevels.HoverMouse();
  sublevels.HoverMouse();
  sublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  TextUtils.writeLog("Navigating to Sub Level");
//  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  var gblSublevels = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
                  
  Sys.HighlightObject(gblSublevels); 
  gblSublevels.HoverMouse();
  gblSublevels.HoverMouse();
  gblSublevels.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  var activebrand = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Company Brands");
  Sys.HighlightObject(activebrand); 
  activebrand.HoverMouse();
  activebrand.HoverMouse();
  activebrand.Click();
//  aqUtils.Delay(3000, Indicator.Text);
  TextUtils.writeLog("Active Company Brand is selected");
//  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
//  var productNmae = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  var productNmae = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")
  Sys.HighlightObject(productNmae); 
  productNmae.HoverMouse();
  productNmae.HoverMouse();
  productNmae.Click();
  productNmae.Keys(brand);
  Sys.HighlightObject(table);
//  aqUtils.Delay(4000, Indicator.Text);
    
  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 52);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 52);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 71);
  ReportUtils.logStep_Screenshot();  
  table.Click(49, 71);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 90);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 90);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==brand){
  table.HoverMouse(49, 109);
  ReportUtils.logStep_Screenshot();
  table.Click(49, 109);
  ValidationUtils.verify(true,true,"Company Brand is available in maconomy");
  }
  
  aqUtils.Delay(3000, "Findind Information");

  TextUtils.writeLog("Company Product is available in maconomy");
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
   


// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
// var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
// }
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}

      //var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl3;
//      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
      ClientApproval.Click();
      //var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      ReportUtils.logStep_Screenshot();
    //var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
//    var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
}


function goToCompanyClient(){ 
  var GblClient = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.TabFolderPanel.TabControl;
  Sys.HighlightObject(GblClient); 
  GblClient.HoverMouse();
  GblClient.HoverMouse();  
  GblClient.Click();
  var active = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  Sys.HighlightObject(active); 
  active.HoverMouse();
  active.HoverMouse(); 
  active.Click();
//  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.ReadPanel.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  Sys.HighlightObject(table); 
  table.HoverMouse();
  table.HoverMouse();
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
}

function CompanyClient(){ 
  //var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.PurchaseApprovalTab
  Sys.HighlightObject(information); 
  information.HoverMouse();
  information.HoverMouse();
  information.Click();
  aqUtils.Delay(3000, "Findind Client Approval");
   


// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
// var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl
 var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl
 Sys.HighlightObject(ClientApproval);
 ClientApproval.HoverMouse();
 ClientApproval.Click();
// }
 if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}

      //var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl3;
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl
      //var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
      ClientApproval.Click();
      //var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      var ApproverTable =  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
      Sys.HighlightObject(ApproverTable);
      ReportUtils.logStep_Screenshot();
    //var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
    //var CloseBar = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel2.TabControl
    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel2.TabControl
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
}
  
function GCData_1_Address(){ 
  aqUtils.Delay(4000, Indicator.Text);;
Sys.Process("Maconomy").Refresh();
//var clientName_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim();
//if(clientName_1!="Full Name")
//ValidationUtils.verify(false,true,"Full Name field is missing in Maconomy for Client Creation");
//else
//ValidationUtils.verify(true,true,"Full Name field is available in Maconomy for Client Creation");
//var strt1_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//if(strt1_1!="Gender")
//ValidationUtils.verify(false,true,"Street 1 field is missing in Maconomy for Client Creation");
//else
//ValidationUtils.verify(true,true,"Street 1 field is available in Maconomy for Client Creation");
//var strt2_1 = Sys.Process("Maconomy").SWTObject("Shell", "New Employee").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
//if(strt2_1!="Country")
//ValidationUtils.verify(false,true,"Street 2 field is missing in Maconomy for Client Creation");
//else
//ValidationUtils.verify(true,true,"Street 2 field is available in Maconomy for Client Creation");
var country_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(country_1!="Country")
ValidationUtils.verify(false,true,"Country field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Country field is available in Maconomy for Client Creation");
var clientlan_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(clientlan_1!="Language")
ValidationUtils.verify(false,true,"Language field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Language field is available in Maconomy for Client Creation");
var taxcode_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(taxcode_1!="Tax No.")
ValidationUtils.verify(false,true,"Tax No. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Tax No. field is available in Maconomy for Client Creation");
var companyReg_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(companyReg_1!="Company Reg. No.")
ValidationUtils.verify(false,true,"Company Reg. No. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Company Reg. No. field is available in Maconomy for Client Creation");
var currency_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(currency_1!="Currency")
ValidationUtils.verify(false,true,"Currency field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Currency field is available in Maconomy for Client Creation");
var clientgrp_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(clientgrp_1!="Client Group")
ValidationUtils.verify(false,true,"Client Group field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Client Group field is available in Maconomy for Client Creation");
var controlAct_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(controlAct_1!="Control Account")
ValidationUtils.verify(false,true,"Control Account field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Control Account field is available in Maconomy for Client Creation");
var bfc_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(bfc_1!="Counter Party BFC")
ValidationUtils.verify(false,true,"Counter Party BFC field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Counter Party BFC field is available in Maconomy for Client Creation");
var madaCode_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(madaCode_1!="Mada Code")
ValidationUtils.verify(false,true,"Mada Code field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Mada Code field is available in Maconomy for Client Creation");
var parentClient_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(parentClient_1!="Parent Client")
ValidationUtils.verify(false,true,"Parent Client field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Parent Client field is available in Maconomy for Client Creation");

}
  
  
function Global_client_Data_1(){ 
GCData_1_Address();
ReportUtils.logStep("INFO","Entering data in Global Client Data 1/1")
var Client_name = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
  
   if(clientName!=""){
 Client_name.setText(clientName+" "+STIME);
 ValidationUtils.verify(true,true,"Client Name Entered in Global Client Data 1/2");
     }
     
var street1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
   if(strt1!=""){
 street1.setText(strt1);
 ValidationUtils.verify(true,true,"Street1 Entered in Global Client Data 1/2");
     }
     
var street2 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
   if(strt2!=""){
 street2.setText(strt1);
  ValidationUtils.verify(true,true,"Street2 Entered in Global Client Data 1/2");
     }
  
//street2.Keys(GCD1[2]);

//var street3 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
//   if(GCD1[3]!=""){
// street3.setText(GCD1[3]);
//     }
////street3.Keys(GCD1[3]);
//var Area = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 2);
//   if(GCD1[4]!=""){
// Area.setText(GCD1[4]);
//     }
////Area.Keys(GCD1[4]);
//var Postal_code = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
//   if(GCD1[5]!=""){
// Postal_code.setText(GCD1[5]);
//     }
////Postal_code.Keys(GCD1[5]);
//aqUtils.Delay(2000, Indicator.Text);;
//var Postal_District = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
//   if(GCD1[6]!=""){
// Postal_District.setText(GCD1[6]);
//     }
////Postal_District.Keys(GCD1[6]);

var Country = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
  if(country!=""){
  Country.Click();
  WorkspaceUtils.DropDownList(country,"Country")
  }
aqUtils.Delay(2000, Indicator.Text);  

var language = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McPopupPickerWidget", "", 2);
  if(clientlan!=""){
  language.Click();
  WorkspaceUtils.DropDownList(clientlan,"Language")
  }
aqUtils.Delay(2000, Indicator.Text);


var Tax_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McTextWidget", "", 2);
   if(taxcode!=""){
 Tax_No.setText(taxcode);
 ValidationUtils.verify(true,true,"Tax No Entered in Global Client Data 1/2");
     }
     
var Compy_Reg_no = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
   if(companyReg!=""){
 Compy_Reg_no.setText(companyReg);
 ValidationUtils.verify(true,true,"Company Registration No Entered in Global Client Data 1/2");
     }
  
var Currency = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2);
  if(currency!=""){
  Currency.Click();
  WorkspaceUtils.DropDownList(currency,"Currency")
  }
aqUtils.Delay(2000, Indicator.Text);  
var client_grp = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2);
  if(clientgrp!=""){
  client_grp.Click();
  WorkspaceUtils.DropDownList(clientgrp,"Client Group")
  }
aqUtils.Delay(2000, Indicator.Text);
var control_Acc = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McPopupPickerWidget", "", 2);
  if(controlAct!=""){
  control_Acc.Click();
  WorkspaceUtils.DropDownList(controlAct,"Control Account")
  }
aqUtils.Delay(2000, Indicator.Text);

var party_BFC = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McValuePickerWidget", "", 2);
if(bfc!=""){
  party_BFC.Click();
  WorkspaceUtils.SearchByValue(party_BFC,"Counter Party BFC",bfc,"Counter Party BFC");
    }

var moda_code = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McValuePickerWidget", "", 2);
//moda_code.Click();
if(madaCode!=""){
  moda_code.Click();
  WorkspaceUtils.SearchByValue(moda_code,"Option",madaCode,"Moda Code");
    }

var parent_client = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McValuePickerWidget", "", 2);
if(parentClient!=""){
  parent_client.Click();
  WorkspaceUtils.SearchByValue(parent_client,"Option",parentClient,"Parent Client");
  }


// if(GCD1[17]!=""){   
////var invoice_spc_Add = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McPopupPickerWidget", "", 2);
// var invoice_spc_Add = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McPopupPickerWidget", "", 2); 
//if(invoice_spc_Add.getText()!=GCD1[17]){
// invoice_spc_Add.Click();
// Sys.Process("Maconomy").Refresh();
//var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
//var Add_Visible7 = true;
//while(Add_Visible7){
//if(list.isEnabled()){
//Add_Visible7 = false;
//    for(var i=list.getItemCount()-1;i>=0;i--){ 
//      if(list.getItem(i).getText_2(0)!=null){ 
//        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==GCD1[17]){ 
//          list.Keys("[Enter]");
//
//          aqUtils.Delay(5000, Indicator.Text);;
//          break;
//        }else{ 
//          list.Keys("[Up]");
//        }
//          
//      }else{ 
//        list.Keys("[Up]");
//      }
//    }
//}
//}
//}
//  aqUtils.Delay(5000, Indicator.Text);;
//}



var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
next.HoverMouse();
ReportUtils.logStep_Screenshot();
next.Click();
}


function GCData_2_Address(){ 
    aqUtils.Delay(4000, Indicator.Text);;
Sys.Process("Maconomy").Refresh();

var company_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(company_1!="Company No.")
ValidationUtils.verify(false,true,"Company No. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Company No. field is available in Maconomy for Client Creation");
var attn_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(attn_1!="Attn.")
ValidationUtils.verify(false,true,"Attn. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Attn. field is available in Maconomy for Client Creation");
var mail_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(mail_1!="E-mail")
ValidationUtils.verify(false,true,"E-mail field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"E-mail field is available in Maconomy for Client Creation");
var phone_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(phone_1!="Phone")
ValidationUtils.verify(false,true,"Phone field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Phone field is available in Maconomy for Client Creation");
var AccDir_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(AccDir_1!="Acct. Director No.")
ValidationUtils.verify(false,true,"Acct. Director No. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Acct. Director No. field is available in Maconomy for Client Creation");
var AccMan_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(AccMan_1!="Account Manager No.")
ValidationUtils.verify(false,true,"Account Manager No. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Account Manager No. field is available in Maconomy for Client Creation");
var Paymentmode_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Paymentmode_1!="Client Payment Mode")
ValidationUtils.verify(false,true,"Client Payment Mode field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Client Payment Mode field is available in Maconomy for Client Creation");
var payterm_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(payterm_1!="Payment Terms")
ValidationUtils.verify(false,true,"Payment Terms field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Payment Terms field is available in Maconomy for Client Creation");
var Comtaxcode_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(Comtaxcode_1!="Company Tax Code")
ValidationUtils.verify(false,true,"Company Tax Code field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Company Tax Code field is available in Maconomy for Client Creation");
var level1Tax_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(level1Tax_1!="Level 1 Tax Derivation")
ValidationUtils.verify(false,true,"Level 1 Tax Derivation field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Level 1 Tax Derivation field is available in Maconomy for Client Creation");
var sales_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(sales_1!="Job Price List, Sales")
ValidationUtils.verify(false,true,"Job Price List, Sales field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Job Price List, Sales field is available in Maconomy for Client Creation");

var intercomp_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(intercomp_1!="Job Price List, Intercomp.")
ValidationUtils.verify(false,true,"Job Price List, Intercomp. field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Job Price List, Intercomp. field is available in Maconomy for Client Creation");

var cost_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(cost_1!="Job Price List, Cost")
ValidationUtils.verify(false,true,"Job Price List, Cost field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Job Price List, Cost field is available in Maconomy for Client Creation");

var standSales_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McTextWidget", "", 1)
standSales_1.Click();
standSales_1 = standSales_1.getText().OleValue.toString().trim() 
if(standSales_1!="Job Price List, Standard Sales")
ValidationUtils.verify(false,true,"Job Price List, Standard Sales field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Job Price List, Standard Sales field is available in Maconomy for Client Creation");

var brand_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 21).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(brand_1!="Default Brand")
ValidationUtils.verify(false,true,"Default Brand field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Default Brand field is available in Maconomy for Client Creation");

var product_1 = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 22).SWTObject("McTextWidget", "", 1).getText().OleValue.toString().trim()
if(product_1!="Default Product")
ValidationUtils.verify(false,true,"Default Product field is missing in Maconomy for Client Creation");
else
ValidationUtils.verify(true,true,"Default Product field is available in Maconomy for Client Creation");

}


function Global_client_Data_2(){
GCData_2_Address();
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(3000, Indicator.Text);
var Company_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
if(company!=""){
  Company_No.Click();
  WorkspaceUtils.SearchByValue(Company_No,"Company",company,"Company Number");
    }



var Attn = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
if(attn!=""){  

Attn.Click();

  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.Keys("[Tab]");
    code.setText(attn);
    code.Keys("[Down]");
//    
//    code.setText("*india*");
    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(7000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
//    Log.Message(table.getItemCount());
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==attn){ 
        Sys.Desktop.KeyDown(0x28); // Down Arrow
        aqUtils.Delay(1000, Indicator.Text);;
        Sys.Desktop.KeyUp(0x28); 
        Sys.Desktop.KeyDown(0x0D);
        Sys.Desktop.KeyUp(0x0D);
        ValidationUtils.verify(true,true,"Attn is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          Attn.setText("");
        }
      }
      
      }
      Sys.Desktop.KeyUp(0x28);
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Contact Person").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
          Attn.setText("");
    }
}

var Email = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);  
if(mail!=""){
 Email.setText(mail);
 ValidationUtils.verify(true,true,"Email Entered in Global Client Data 2/2");
     }

var Phone = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2); 
if(phone!=""){
 Phone.setText(phone);
 ValidationUtils.verify(true,true,"Phone Number is Entered in Global Client Data 2/2");
     }

var Acct_Director_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("McValuePickerWidget", "", 2); 
if(AccDir!=""){
  Acct_Director_No.Click();
  WorkspaceUtils.SearchByValueTable(Acct_Director_No,"Employee",AccDir,"Acct Director No");
}



var Account_Manager_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McValuePickerWidget", "", 2); 
if(AccMan!=""){
  Account_Manager_No.Click();
  WorkspaceUtils.SearchByValueTable(Account_Manager_No,"Employee",AccMan,"Acct Manager No");
}

var Client_Payment_Mode = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McValuePickerWidget", "", 2); 
if(Paymentmode!=""){
  Client_Payment_Mode.Click();
  WorkspaceUtils.SearchByValue(Client_Payment_Mode,"Client Payment Mode",Paymentmode,"Client Payment Mode");
         }


var Payment_Terms = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 12).SWTObject("McPopupPickerWidget", "", 2); 
if(payterm !=""){
  Payment_Terms.Click();  
  aqUtils.Delay(5000, Indicator.Text);;
  Sys.Process("Maconomy").Refresh(); 
  WorkspaceUtils.DropDownList(payterm,"Payment Terms"); 
     }


var Company_Tax_Code = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2); 
if(Comtaxcode!=""){
  Company_Tax_Code.Click();  
  aqUtils.Delay(5000, Indicator.Text);;
  WorkspaceUtils.DropDownList(Comtaxcode,"Company Tax Code"); 
     }


var Level_1_Tax_Derivation = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McValuePickerWidget", "", 2); 
Level_1_Tax_Derivation.setText("-");




var Job_Price_List_Sales = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McValuePickerWidget", "", 2); 
if(sales!=""){
  Job_Price_List_Sales.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Sales,"Job Price List",sales,"Job Price List Sales");
         }


var Job_Price_List_Intercomp = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 18).SWTObject("McValuePickerWidget", "", 2); 
if(intercomp!=""){
  Job_Price_List_Intercomp.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Intercomp,"Job Price List",intercomp,"Job Price List Intercomp");
         }

var Job_Price_List_Cost = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McValuePickerWidget", "", 2); 
if(cost!=""){
  Job_Price_List_Cost.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Cost,"Job Price List",cost,"Job Price List Cost");
         }
  
var Job_Price_List_Standard_Sales = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McValuePickerWidget", "", 2); 
if(standSales!=""){
  Job_Price_List_Standard_Sales.Click();
  WorkspaceUtils.SearchByValue(Job_Price_List_Standard_Sales,"Job Price List",standSales,"Job Price List Standard Sales");
         }

var Default_Brand = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 21).SWTObject("McTextWidget", "", 2);
if(brand!=""){
  Default_Brand.setText(brand+" "+STIME);
         }
  
var Default_Product = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 22).SWTObject("McTextWidget", "", 2);
if(product!=""){
  Default_Product.setText(product+" "+STIME);
         }
   
var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
Sys.HighlightObject(next); 
next.HoverMouse();
ReportUtils.logStep_Screenshot();
next.Click();

aqUtils.Delay(3000, Indicator.Text);;

var label = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();

aqUtils.Delay(5000, Indicator.Text);

//Sys.Process("Maconomy").Refresh();
//var mainparent = Sys.Process("Maconomy")
//aqUtils.Delay(3000, "Waiting to find Object");
//var childCount = Sys.Process("Maconomy").ChildCount;
//for(var ci=0;ci<childCount;ci++){ 
//  if((mainparent.Child(ci).Name!="JavaRuntime()")&&(mainparent.Child(ci).Visible!=false)){
//var  Full_Name = mainparent.Child(ci).WndCaption.toString().trim();
//Log.Message(Full_Name);
//if(Full_Name.indexOf("Client Management - Client Information Card")!=-1){
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//ReportUtils.logStep("INFO",lab)
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//  Ok.Click();
//  break;
//}
//}
//}

//----------------This Case for same Company Registration Number can use for more than one Client
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Client Management - Client Information Card"){
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//ReportUtils.logStep("INFO",lab)
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Client Information Card").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//  Ok.Click();
//
//}


}


function gotoCreatedClient(){ 
  aqUtils.Delay(4000, Indicator.Text);;
var inactive = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Inactive Customers");
inactive.HoverMouse();
ReportUtils.logStep_Screenshot();
inactive.Click();
aqUtils.Delay(5000, Indicator.Text);;
var ClientNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1)
ClientNo.Keys("[Tab][Tab]")
var ClientName = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
ClientName.Keys(clientName+" "+STIME);
aqUtils.Delay(5000, Indicator.Text);;

var cltTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var flag=false;
  for(var v=0;v<cltTable.getItemCount();v++){ 
    if(cltTable.getItem(v).getText_2(2).OleValue.toString().trim()==(clientName+" "+STIME)){ 

      flag=true;
      ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Client Number",EnvParams.Opco,"Data Management",cltTable.getItem(v).getText_2(0).OleValue.toString().trim())
      break;
    }else{ 
    cltTable.Keys("[Down]");
  }      
  }
ClientNumber =  cltTable.getItem(v).getText_2(0).OleValue.toString().trim(); 
  ValidationUtils.verify(flag,true,"Client Created is available in system");
  ValidationUtils.verify(true,true,"Client Number :"+cltTable.getItem(v).getText_2(0).OleValue.toString().trim());
  
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(3000, Indicator.Text);;
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
if(ImageRepository.ImageSet.Forward.Exists()){ 

}else{
var approveAction = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
approveAction.Click();
aqUtils.Delay(3000, Indicator.Text);;
}
var document = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
document.HoverMouse();
ReportUtils.logStep_Screenshot();
document.Click();
aqUtils.Delay(3000, Indicator.Text);;

Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var atthDoc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
atthDoc.HoverMouse();
ReportUtils.logStep_Screenshot();
atthDoc.Click();
aqUtils.Delay(4000, Indicator.Text);;
var dicratory = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
dicratory.Keys(workBook);
var opendoc = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("Button", "&Open", 1);
Sys.HighlightObject(opendoc);
opendoc.HoverMouse();
ReportUtils.logStep_Screenshot();
opendoc.Click();


var home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//home.HoverMouse();
//ReportUtils.logStep_Screenshot();
home.Click();
aqUtils.Delay(3000, Indicator.Text);;
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
info.HoverMouse();
ReportUtils.logStep_Screenshot();
info.Click();
aqUtils.Delay(5000, Indicator.Text);;
var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7)
Sys.HighlightObject(submit);
submit.HoverMouse();
ReportUtils.logStep_Screenshot();
submit.Click();
aqUtils.Delay(8000, Indicator.Text);

//--------------This case is used for when same company Registration Number can be used for more Client----------
//if(ImageRepository.ImageSet.OK_Button.Exists()){ 
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Information").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//ReportUtils.logStep("INFO",lab)
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse();
//ReportUtils.logStep_Screenshot("");
//  Ok.Click();
//}
//
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Client Management - Information"){
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Information").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Client Management - Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse();   
//    Ok.Click();
//}

aqUtils.Delay(2000, Indicator.Text);;
if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}




aqUtils.Delay(3000, Indicator.Text);;
var AllApproved = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 9);
AllApproved.HoverMouse();
ReportUtils.logStep_Screenshot();
AllApproved.Click();
aqUtils.Delay(4000, Indicator.Text);;
var y =0 ;
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var ApproverTable = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.ClientApproval.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
//Log.Message(ApproverTable.FullName);
//var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
for(var i=0;i<ApproverTable.getItemCount();i++){ 
  
   var approvers="";
       approvers = EnvParams.Opco+"*"+ClientNumber+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +i+ ": " +approvers);
Approve_Level[y] = approvers;
       y++;
}
var moreinfo = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
//Log.Message(moreinfo.FullName);
moreinfo.Click();
//Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "").Click();
aqUtils.Delay(3000, Indicator.Text);;
if(ImageRepository.ImageSet.Forward.Exists()){ 
  ImageRepository.ImageSet.Forward.Click();
}

aqUtils.Delay(4000, Indicator.Text);;
var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
level = 1;
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();  
  aqUtils.Delay(8000, Indicator.Text);; 
}
}
}



//function todo(lvl,clientLvl){ 
//    var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
//  toDo.DBlClick();
//  aqUtils.Delay(3000, Indicator.Text);
//
//  //To Maximaize the window
//  Sys.Desktop.KeyDown(0x12);
//  Sys.Desktop.KeyDown(0x20);
//  Sys.Desktop.KeyUp(0x12);
//  Sys.Desktop.KeyUp(0x20);
//  Sys.Desktop.KeyDown(0x58);
//  Sys.Desktop.KeyUp(0x58);  
//  aqUtils.Delay(1000, Indicator.Text);
//
//  
//  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
//  var refresh;
//for(var i=1;i<=childCC;i++){ 
//refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
//if(refresh.isVisible()){ 
//refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
//refresh.Click();
//
//  
//  
//  aqUtils.Delay(15000, Indicator.Text);
//Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
//if(Client_Managt.isVisible()){ 
//Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
//if(clientLvl==(ApproveInfo.length-1)){
//if(lvl==3){
//Client_Managt.ClickItem("|Approve Customer by Type (Substitute) (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Customer by Type (Substitute) (*)");
//}
//if(lvl==2){
//Client_Managt.ClickItem("|Approve Customer by Type (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Customer by Type (*)");
//}
//}
//else{ 
// if(lvl==3){
//Client_Managt.ClickItem("|Approve Customer (Substitute) (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Customer (Substitute) (*)");
//}
//if(lvl==2){
//Client_Managt.ClickItem("|Approve Customer (*)");
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|Approve Customer (*)");
//} 
//}
//break;
//}
//}
//}
//}


//function CredentialLogin(){ 
//for(var i=level;i<Approve_Level.length;i++){
//  var UserN = true;
//  var temp="";
//  var Cred = Approve_Level[i].split("*");
//  for(var j=2;j<4;j++){
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
//  { 
//
//     var sheetName = "Agency Users";
//     workBook = Project.Path+excelName;
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//  }
//  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
//  { 
//
//    var sheetName = "SSC Users";
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
//  }
//  else{ 
//   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
//    if(UserN){ 
//      goToHR();
//      UserN = false;
//    }
//    temp = searchNumber(Eno);
//  }
////  Log.Message(temp)
//  if(temp.length!=0){
//    temp = temp+"*"+j;
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
//  break;
//  }
//  }
//  if((temp=="")||(temp==null))
//  Log.Error("User Name is Not available for level :"+i);
////  Log.Message("Logins :"+temp);
//}
//WorkspaceUtils.closeAllWorkspaces();
//
//}


//function FinalApproveClient(comID,cltID,apvr,clientLvl){ 
////function FinalApproveClient(){
////  var ApproveInfo =[];
////  ApproveInfo[0] = "7";
////  ApproveInfo[1] = "7";
////  var comID ="1307";
////  var cltID = "111314";
////  var apvr ="CHFP CT Clients";
////  var clientLvl ="1";
//  aqUtils.Delay(5000, Indicator.Text);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){
//aqUtils.Delay(2000, Indicator.Text);
//ImageRepository.ImageSet.Show_Filter.Click();
//}
//
//
//var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
//.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
//.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//firstCell.setText(cltID);
////firstCell.Keys("[Tab]");
//var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//  
////var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
////  job.setText(empNumber)
////job.setText(FullName + " "+STIME);
//aqUtils.Delay(6000, Indicator.Text);;
//var flag=false;
//for(var v=0;v<table.getItemCount();v++){ 
//  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==cltID){ 
////if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){
//
//    flag=true;    
//    break;
//  }
//  else{ 
//    table.Keys("[Down]");
//  }
//}
//    
//    
//    
//    
//
//ValidationUtils.verify(flag,true,"Created Client is available in system");
//if(flag){ 
//closefilter.HoverMouse();
//ReportUtils.logStep_Screenshot();
//closefilter.Click();
//aqUtils.Delay(5000, Indicator.Text);;
//var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
//Sys.HighlightObject(Approve)
//if(Approve.isEnabled()){ 
//Approve.HoverMouse();
//ReportUtils.logStep_Screenshot();
//  Approve.Click();
//  if(clientLvl==(ApproveInfo.length-1)){
// for(var j=0;j<12;j++){ 
//    if(ImageRepository.ImageSet.Ok.Exists()){ 
////     ImageRepository.ImageSet.Ok.Click();
////     aqUtils.Delay(1000, Indicator.Text);;
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse(); 
//    ReportUtils.logStep_Screenshot();
//    Ok.Click(); 
//    aqUtils.Delay(8000, Indicator.Text); ;
//   }
//   else if(ImageRepository.ImageSet.OK_Button.Exists()){ 
////     ImageRepository.ImageSet.OK_Button.Click();
////     aqUtils.Delay(1000, Indicator.Text);;
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse(); 
//    ReportUtils.logStep_Screenshot();
//    Ok.Click(); 
//    aqUtils.Delay(8000, Indicator.Text); ;
//   }
// }
//    for(var j=0;j<12;j++){ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Approve Customer by Type"){
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//    Ok.HoverMouse(); 
//    ReportUtils.logStep_Screenshot();
//    Ok.Click(); 
//    aqUtils.Delay(8000, Indicator.Text); ;
//}
//}
//      
//}
// aqUtils.Delay(2000, Indicator.Text);;
//if(ImageRepository.ImageSet.Maximize.Exists()){
//ImageRepository.ImageSet.Maximize.Click();
//}else{ 
//var sideBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
//sideBar.Click();
// aqUtils.Delay(2000, Indicator.Text);;
// ImageRepository.ImageSet.Maximize.Click();
//}
//var AllApproved = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.CTClientApproval;
//AllApproved.Click();
//aqUtils.Delay(8000, Indicator.Text); ;
//ReportUtils.logStep_Screenshot();
//var closeInfor = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
//closeInfor.Click();
//var showfilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
//showfilter.Click();
//aqUtils.Delay(5000, Indicator.Text); ;
//var activeCustomer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Customers")
//activeCustomer.Click();
//var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
//.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
//.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//firstCell.Click();
//firstCell.setText("^a[BS]");
//firstCell.setText(cltID);
//var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//aqUtils.Delay(6000, Indicator.Text);;
//var flag=false;
//if(table.getItemCount()==3){ 
//  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==cltID){ 
//  aqUtils.Delay(8000, Indicator.Text);;
//  ReportUtils.logStep_Screenshot();
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//      ExcelUtils.WriteExcelSheet("Brand Number",EnvParams.Opco,"Data Management",table.getItem(1).getText_2(0).OleValue.toString().trim())
//      ExcelUtils.WriteExcelSheet("Product Number",EnvParams.Opco,"Data Management",table.getItem(2).getText_2(0).OleValue.toString().trim())
//  }
//
//}
//aqUtils.Delay(2000, Indicator.Text);;
//  }
//  }
//
//
//}

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
