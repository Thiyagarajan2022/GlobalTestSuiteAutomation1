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
var sheetName = "CreateGlobalProduct";
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
function CreateGlobalProduct(){
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
sheetName = "CreateGlobalProduct";
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

gotoMenu(); 
gotoClientSearch();
globalClient(); 
SubLevels();
GlobalClient_Screen2();
PopUp();
SelectProduct();
if(EnvParams.Country.toUpperCase()=="INDIA"){
Runner.CallMethod("IND_CreateGlobalBrand.indiaSpecific",State,GST,PAN,TAN,TIN);
}
if(EnvParams.Country.toUpperCase()=="SPAIN"){
Runner.CallMethod("SPA_GlobalBrand.spainSpecific",SII_Tax);
}
Information();
ApprvalInformation();

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

//ExcelUtils.setExcelName(workBook, sheetName, true);
//brandname = ExcelUtils.getRowDatas("Brand Name",EnvParams.Opco)
//
//if((brandname==null)||(brandname=="")){ 
//ValidationUtils.verify(false,true,"Brand Name is Needed to Create a Client");
//}


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



function SubLevels(){ 
  
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
  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All").OleValue.toString().trim());;
  waitForObj(activeBrand);
  activeBrand.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, "Playback");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var newglbbrand= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 4);
  waitForObj(newglbbrand);
  if(newglbbrand.isEnabled()){ 
   newglbbrand.Click(); 
  }else{ 
    
  if(newglbbrand.isEnabled()){ 
   newglbbrand.Click(); 
  }else{ 
  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active").OleValue.toString().trim());;
  waitForObj(activeBrand);
  activeBrand.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, "Playback");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  } 

  if(newglbbrand.isEnabled()){ 
   newglbbrand.Click(); 
  }else{ 
  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Brands").OleValue.toString().trim());;
  waitForObj(activeBrand);
  activeBrand.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, "Playback");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }

  if(newglbbrand.isEnabled()){ 
   newglbbrand.Click(); 
  }else{ 
  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Active Products").OleValue.toString().trim());;
  waitForObj(activeBrand);
  activeBrand.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(5000, "Playback");
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
  waitForObj(newglbbrand);
  newglbbrand.Click(); 
   
  }
  
  }
  
  }

  }
  
 Sys.HighlightObject(newglbbrand);
 aqUtils.Delay(8000, "Playback");
var cancel = Sys.Process("Maconomy").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
//   var cancel = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
 waitForObj(cancel)
 
//var brandname1=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
//brandname1.Click();
//Sys.HighlightObject(brandname1);
//brandname1.setText(brandname+" "+STIME);
//brandname = brandname+" "+STIME;

var defaultname1=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
 // waitForObj(brandname1);
  defaultname1.Click();
Sys.HighlightObject(defaultname1);
//defaultname1.setText(defaultname);
 defaultname1.setText(defaultname+" "+STIME);
 brandname = defaultname+" "+STIME;
 
//Default Name
 
var next=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.Composite.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "&Next >").OleValue.toString().trim())
 next.HoverMouse();
 Sys.HighlightObject(next);
 ReportUtils.logStep_Screenshot() ;
 next.Click();
 
//Global Client - Client Information Card
}



function PopUp()
{
  
var w = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Global Client - Client Information Card").OleValue.toString().trim());
var Label = w.SWTObject("Label", "*").getText();
ReportUtils.logStep("INFO",Label.OleValue.toString().trim());
var OK = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
OK.Click();

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

  aqUtils.Delay(5000, "Finding pop-up");

     
    
 if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Global Client - Client Information Card")    
    { 
  var button1=Aliases.Maconomy.CreateGlobalBrand3.Composite.Button;
  var label =Aliases.Maconomy.CreateGlobalBrand3.SWTObject("Label", "*").WndCaption;
  Log.Message(label );
       button1.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      button1.Click();
      aqUtils.Delay(5000, "Finding pop-up");
}
}



function SelectProduct(){

if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
  
}
  var activeBrand = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Blocked").OleValue.toString().trim());;
  waitForObj(activeBrand);
  activeBrand.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  
var name1=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
name1.Click();
  name1.Keys(brandname); 
var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 Sys.HighlightObject(table);

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
sheetName = "CreateGlobalProduct";
if(OpCo2[2]==Project_manager){
level = 1;

//if((EnvParams.Country.toUpperCase()=="INDIA")||(EnvParams.Country.toUpperCase()=="SPAIN")){
//var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite2
//}else{
////var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite2.SingleToolItemControl
//var Approve = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite;
//}
//
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
ValidationUtils.verify(true,true,"Global Product is Approved by "+Project_manager)
TextUtils.writeLog("Levels 0 has  Approved the Created Global Product");
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
//    ApproveInfo[0] = Cred[0]+"*"+Cred[1]+"*"+"1707 Management (TST)*3"
//    ApproveInfo[1] = Cred[0]+"*"+Cred[1]+"*"+"1707 Senior Finance (TST)*2"
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
var Client_Managt=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;

//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
//Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
//}
//if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
//Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
//}
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
var Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel;
for(var j=0;j<Approve.ChildCount;j++){ 
 if(Approve.Child(j).isVisible()){ 
   Approve = Approve.Child(j);
 for(var i=0;i<Approve.ChildCount;i++){ 
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
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
  if((Approve.Child(i).isVisible())&&(Approve.Child(i).Name.indexOf("SingleToolItemControl")!=-1)&&(Approve.Child(i).text==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve").OleValue.toString().trim())){
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
aqUtils.Delay(8000, Indicator.Text);;
TextUtils.writeLog("Global Client is Approved by "+Apvr);
if(Approve_Level.length==lvl+1){
//var label = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Customer by Type").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
//Ok.HoverMouse(); 
//ReportUtils.logStep_Screenshot();
//aqUtils.Delay(4000, Indicator.Text); ;
//Ok.Click(); 
aqUtils.Delay(20000, Indicator.Text); ;
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
  ExcelUtils.WriteExcelSheet("Global Product No",EnvParams.Opco,"Data Management",ClientNum)
  ExcelUtils.WriteExcelSheet("GlobalProduct Name",EnvParams.Opco,"Data Management",GBName)
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
