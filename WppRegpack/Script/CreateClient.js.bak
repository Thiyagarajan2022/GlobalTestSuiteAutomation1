//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

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
var clientName,strt1,strt2,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,madaCode,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product ="";
var ClientNumber = "";

function ClientCreation(){ 
Indicator.PushText("waiting for window to open");
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
Restart.login(Project_manager);
  
}

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreateClient";
Language = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
clientName,strt1,strt2,country,clientlan,taxcode,companyReg,currency,clientgrp,controlAct,bfc,madaCode,parentClient,ISA,company,attn,mail,phone,AccDir,AccMan,Paymentmode,payterm,Comtaxcode,level1Tax,sales,intercomp,cost,standSales,brand,product ="";
ClientNumber = "";
Approve_Level = [];

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Log.Message(EnvParams.Opco)
Log.Message(Language)
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Client Creation started::"+STIME);
getDetails();
goToJobMenuItem(); 
selectNewClient(); 
Global_client_Data_1();
Global_client_Data_2();
gotoCreatedClient();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3],i);
FinalApproveClient(temp[0],temp[1],temp[2],i);
}
//FinalApproveClient();
}


function goToJobMenuItem(){
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
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|Client Management");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|Client Management");
}

} 

aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receviable Menu");

}

function getDetails(){ 
  ExcelUtils.setExcelName(workBook, sheetName, true);
clientName = ExcelUtils.getRowDatas("Client Name",EnvParams.Opco)
if((clientName==null)||(clientName=="")){ 
ValidationUtils.verify(false,true,"Client Name is Needed to Create a Client");
}

strt1 = ExcelUtils.getRowDatas("Street 1",EnvParams.Opco)
if((strt1==null)||(strt1=="")){ 
ValidationUtils.verify(false,true,"Street 1 is Needed to Create a Client");
}

strt2 = ExcelUtils.getRowDatas("Street 2",EnvParams.Opco)
if((strt2==null)||(strt2=="")){ 
ValidationUtils.verify(false,true,"Street 2 is Needed to Create a Client");
}

country = ExcelUtils.getRowDatas("Country",EnvParams.Opco)
if((country==null)||(country=="")){ 
ValidationUtils.verify(false,true,"Country is Needed to Create a Client");
}
clientlan = ExcelUtils.getRowDatas("Language",EnvParams.Opco)
if((clientlan==null)||(clientlan=="")){ 
ValidationUtils.verify(false,true,"Language is Needed to Create a Client");
}
taxcode = ExcelUtils.getRowDatas("Tax No.",EnvParams.Opco)
if((taxcode==null)||(taxcode=="")){ 
ValidationUtils.verify(false,true,"Tax No. is Needed to Create a Client");
}

companyReg = ExcelUtils.getRowDatas("Company Reg. No.",EnvParams.Opco)
if((companyReg==null)||(companyReg=="")){ 
ValidationUtils.verify(false,true,"Company Reg. No. is Needed to Create a Client");
}
currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((currency==null)||(currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Create a Client");
}

clientgrp = ExcelUtils.getRowDatas("Client Group",EnvParams.Opco)
if((clientgrp==null)||(clientgrp=="")){ 
ValidationUtils.verify(false,true,"Client Group is Needed to Create a Client");
}
controlAct = ExcelUtils.getRowDatas("Control Account",EnvParams.Opco)
if((controlAct==null)||(controlAct=="")){ 
ValidationUtils.verify(false,true,"Control Account is Needed to Create a Client");
}

bfc = ExcelUtils.getRowDatas("Counter Party BFC",EnvParams.Opco)
if((bfc==null)||(bfc=="")){ 
ValidationUtils.verify(false,true,"Counter Party BFC is Needed to Create a Client");
}
madaCode = ExcelUtils.getRowDatas("Mada Code",EnvParams.Opco)
if((madaCode==null)||(madaCode=="")){ 
ValidationUtils.verify(false,true,"Mada Code is Needed to Create a Client");
}
parentClient = ExcelUtils.getRowDatas("Parent Client",EnvParams.Opco)
if((parentClient==null)||(parentClient=="")){ 
ValidationUtils.verify(false,true,"Parent Client is Needed to Create a Client");
}

company = ExcelUtils.getRowDatas("Company No.",EnvParams.Opco)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company No. is Needed to Create a Client");
}
attn = ExcelUtils.getRowDatas("Attn.",EnvParams.Opco)
if((attn==null)||(attn=="")){ 
ValidationUtils.verify(false,true,"Attn. is Needed to Create a Client");
}
mail = ExcelUtils.getRowDatas("E-mail",EnvParams.Opco)
if((mail==null)||(mail=="")){ 
ValidationUtils.verify(false,true,"E-mail is Needed to Create a Client");
}
phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
if((phone==null)||(phone=="")){ 
ValidationUtils.verify(false,true,"Phone is Needed to Create a Client");
}
AccDir = ExcelUtils.getRowDatas("Acct. Director No.",EnvParams.Opco)
if((AccDir==null)||(AccDir=="")){ 
ValidationUtils.verify(false,true,"Acct. Director No. is Needed to Create a Client");
}
AccMan = ExcelUtils.getRowDatas("Account Manager No.",EnvParams.Opco)
if((AccMan==null)||(AccMan=="")){ 
ValidationUtils.verify(false,true,"Account Manager No. is Needed to Create a Client");
}
Paymentmode = ExcelUtils.getRowDatas("Client Payment Mode",EnvParams.Opco)
if((Paymentmode==null)||(Paymentmode=="")){ 
ValidationUtils.verify(false,true,"Client Payment Mode is Needed to Create a Client");
}
payterm = ExcelUtils.getRowDatas("Payment Terms",EnvParams.Opco)
if((payterm==null)||(payterm=="")){ 
ValidationUtils.verify(false,true,"Payment Terms is Needed to Create a Client");
}
Comtaxcode = ExcelUtils.getRowDatas("Company Tax Code",EnvParams.Opco)
if((Comtaxcode==null)||(Comtaxcode=="")){ 
ValidationUtils.verify(false,true,"Company Tax Code is Needed to Create a Client");
}

sales = ExcelUtils.getRowDatas("Job Price List, Sales",EnvParams.Opco)
if((sales==null)||(sales=="")){ 
ValidationUtils.verify(false,true,"Job Price List, Sales is Needed to Create a Client");
}
intercomp = ExcelUtils.getRowDatas("Job Price List, Intercomp.",EnvParams.Opco)
if((intercomp==null)||(intercomp=="")){ 
ValidationUtils.verify(false,true,"Job Price List, Intercomp. is Needed to Create a Client");
}
cost = ExcelUtils.getRowDatas("Job Price List, Cost",EnvParams.Opco)
if((cost==null)||(cost=="")){ 
ValidationUtils.verify(false,true,"Job Price List, Cost is Needed to Create a Client");
}
standSales = ExcelUtils.getRowDatas("Job Price List, Standard Sales",EnvParams.Opco)
if((standSales==null)||(standSales=="")){ 
ValidationUtils.verify(false,true,"Job Price List, Standard Sales is Needed to Create a Client");
}
brand = ExcelUtils.getRowDatas("Default Brand",EnvParams.Opco)
if((brand==null)||(brand=="")){ 
ValidationUtils.verify(false,true,"Default Brand is Needed to Create a Client");
}
product = ExcelUtils.getRowDatas("Default Product",EnvParams.Opco)
if((product==null)||(product=="")){ 
ValidationUtils.verify(false,true,"Default Product is Needed to Create a Client");
}

}


function selectNewClient(){

  aqUtils.Delay(6000, Indicator.Text);;

var mparent = false;
mainParent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
var mpChildCount = mainParent.ChildCount;
for(var mp=0;mp<mpChildCount;mp++){ 
  if(mainParent.Child(mp).isVisible()){ 
    var secChild = mainParent.Child(mp);
    var scChildCount = secChild.ChildCount;
    for(var sc=0;sc<scChildCount;sc++){ 
  if((secChild.Child(sc).isVisible())&&(secChild.Child(sc).Name.indexOf("Composite")!=-1)&&(secChild.Child(sc).Name.indexOf("1")!=-1)){ 
    mainParent = secChild.Child(sc);
    mparent = true;
    break;
    }
    }
    if(mparent){ 
      break;
    }
    
  }
}

//Log.Message(mainParent.FullName)

var myClient = mainParent.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "My Clients");
myClient.Click();
aqUtils.Delay(5000, Indicator.Text);
var NewClient = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
NewClient.HoverMouse();
ReportUtils.logStep_Screenshot("");
NewClient.Click();
ReportUtils.logStep("INFO", "Enter Job Details");

var DueDiligence = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 26).SWTObject("McPopupPickerWidget", "", 2);
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(4000, Indicator.Text);
DueDiligence.Keys("Yes");
  aqUtils.Delay(2000, Indicator.Text);
var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
next.HoverMouse();
ReportUtils.logStep_Screenshot("");
next.Click();


var client_identification = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(4000, Indicator.Text);
  client_identification.Keys("Yes");
 
 aqUtils.Delay(3000, Indicator.Text);
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
 checks_did_you_perform.setText("Yes");

  
 var new_client_business = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("McTextWidget", "", 2);
 new_client_business.setText("Yes");

  
 var company_owner = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McPopupPickerWidget", "", 2);
  company_owner.Keys("Yes");


   aqUtils.Delay(2000, Indicator.Text);
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
 checks_did_you_perform.setText("Yes");

  
 var foreign_jurisdictions = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McPopupPickerWidget", "", 2);
  foreign_jurisdictions.Keys("Yes");

    aqUtils.Delay(2000, Indicator.Text);
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 11).SWTObject("McTextWidget", "", 2);
 checks_did_you_perform.setText("Yes");

  
 var sanction_lists = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 13).SWTObject("McPopupPickerWidget", "", 2);
 sanction_lists.Keys("Yes");
  aqUtils.Delay(2000, Indicator.Text);
  
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 14).SWTObject("McTextWidget", "", 2);
 checks_did_you_perform.setText("Yes");

  
 var potential_client_conflicts = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McPopupPickerWidget", "", 2);
  potential_client_conflicts.Keys("Yes");

  
   aqUtils.Delay(2000, Indicator.Text); 
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 17).SWTObject("McTextWidget", "", 2);
 checks_did_you_perform.setText("Yes");

  
 var new_client_can_pay = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 19).SWTObject("McPopupPickerWidget", "", 2);
  new_client_can_pay.Keys("Yes");

   aqUtils.Delay(2000, Indicator.Text); 
 var checks_did_you_perform = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 20).SWTObject("McTextWidget", "", 2);
 checks_did_you_perform.setText("Yes");

  aqUtils.Delay(2000, Indicator.Text);
 var services_provided_new_client = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 23).SWTObject("McTextWidget", "", 2);
 services_provided_new_client.setText("Yes");


  var next = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Button", "&Next >");
  Sys.HighlightObject(next);
  next.HoverMouse();
ReportUtils.logStep_Screenshot("");
next.Click();

  
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
//Acct_Director_No.Click();
//
//  aqUtils.Delay(1000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x11);
//    Sys.Desktop.KeyDown(0x47);
//    Sys.Desktop.KeyUp(0x11);
//    Sys.Desktop.KeyUp(0x47);
//    aqUtils.Delay(3000, Indicator.Text);;
//    var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//    code.setText(AccDir);
//    aqUtils.Delay(3000, Indicator.Text);;
//    code.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
//    ImageRepository.ImageSet.sale_dropDown.Click();
//var code = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);
//    code.Keys("Yes");
//    aqUtils.Delay(2000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x0D);
//    Sys.Desktop.KeyUp(0x0D);
//    aqUtils.Delay(3000, Indicator.Text);;
//    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
//    Sys.HighlightObject(serch);
//    serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
//    var table = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
//    Sys.HighlightObject(table);
//    Log.Message(table.getItemCount());
//    var itemCount = table.getItemCount();
//    if(itemCount>0){ 
//    for(var i=0;i<itemCount;i++){
//    Log.Message("7th Column :"+table.getItem(i).getText_2(7));
//   if (table.getItem(i).getText_2(7)!=null){
//      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==AccDir) { 
//        Sys.Desktop.KeyDown(0x28); // Down Arrow
//        aqUtils.Delay(1000, Indicator.Text);;
//        Sys.Desktop.KeyUp(0x28); 
//        Sys.Desktop.KeyDown(0x0D);
//        Sys.Desktop.KeyUp(0x0D);
//      }
//      else{ 
//        Sys.Desktop.KeyDown(0x28);
//        if(i==itemCount-1){ 
//          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//          cancel.Click();
//          aqUtils.Delay(1000, Indicator.Text);;
//          Acct_Director_No.setText("");
//        }
//      }
//      
//      }
//      else { 
//      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//      cancel.Click();
//      aqUtils.Delay(1000, Indicator.Text);;
//          Acct_Director_No.setText("");
//    }
//      }
//      Sys.Desktop.KeyUp(0x28);
//    }
//
//    else { 
//      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//      cancel.Click();
//      aqUtils.Delay(1000, Indicator.Text);;
//          Acct_Director_No.setText("");
//    }
//     }


var Account_Manager_No = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("McValuePickerWidget", "", 2); 
if(AccMan!=""){
  Account_Manager_No.Click();
  WorkspaceUtils.SearchByValueTable(Account_Manager_No,"Employee",AccMan,"Acct Manager No");
}
//if(AccMan!=""){
//Account_Manager_No.Click();
//
//  aqUtils.Delay(1000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x11);
//    Sys.Desktop.KeyDown(0x47);
//    Sys.Desktop.KeyUp(0x11);
//    Sys.Desktop.KeyUp(0x47);
//    aqUtils.Delay(3000, Indicator.Text);;
//    var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//    code.setText(AccMan);
//    aqUtils.Delay(3000, Indicator.Text);;
//    code.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
//    ImageRepository.ImageSet.sale_dropDown.Click();
//var code = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").SWTObject("Grid", "", 3);
//    code.Keys("Yes");
//    aqUtils.Delay(2000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x0D);
//    Sys.Desktop.KeyUp(0x0D);
//    aqUtils.Delay(3000, Indicator.Text);;
//    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
//    Sys.HighlightObject(serch);
//    serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
//    var table = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
//    Sys.HighlightObject(table);
//    Log.Message(table.getItemCount());
//    var itemCount = table.getItemCount();
//    if(itemCount>0){ 
//    for(var i=0;i<itemCount;i++){
//    Log.Message("7th Column :"+table.getItem(i).getText_2(7));
//   if (table.getItem(i).getText_2(7)!=null){
//      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==AccMan) { 
//        Sys.Desktop.KeyDown(0x28); // Down Arrow
//        aqUtils.Delay(1000, Indicator.Text);;
//        Sys.Desktop.KeyUp(0x28); 
//        Sys.Desktop.KeyDown(0x0D);
//        Sys.Desktop.KeyUp(0x0D);
//      }
//      else{ 
//        Sys.Desktop.KeyDown(0x28);
//        if(i==itemCount-1){ 
//          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//          cancel.Click();
//          aqUtils.Delay(1000, Indicator.Text);;
//          Account_Manager_No.setText("");
//        }
//      }
//      
//      }
//      }
//      Sys.Desktop.KeyUp(0x28);
//    }
//    else { 
//      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
//      cancel.Click();
//      aqUtils.Delay(1000, Indicator.Text);;
//          Account_Manager_No.setText("");
//    }
//         }else{ 
//    ValidationUtils.verify(false,true,"Accountt Director No Needed to create Client in Global Client Data 2/2");
//  }


//var Budget_Holder = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McValuePickerWidget", "", 2); 
////Budget_Holder.Click();
//if(GCD2[7]!=""){
//  Budget_Holder.Click();
//  WorkspaceUtils.SearchByValue(Budget_Holder,"Employee",GCD2[7]);
//  } 
//   
//var Main_Biller = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("McValuePickerWidget", "", 2); 
//if(GCD2[8]!=""){
//  Main_Biller.Click();
//  WorkspaceUtils.SearchByValue(Main_Biller,"Employee",GCD2[8]);
//  }
//var Client_Finance = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 10).SWTObject("McValuePickerWidget", "", 2); 
//if(GCD2[9]!=""){
//  Client_Finance.Click();
//  WorkspaceUtils.SearchByValue(Client_Finance,"Employee",GCD2[9]);
//  }

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


//var Client_Specific_Logo = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McPopupPickerWidget", "", 2); 
//if(GCD2[14]!=""){
//  Client_Specific_Logo.Click();  
//  aqUtils.Delay(5000, Indicator.Text);;
//  Sys.Process("Maconomy").Refresh(); 
//  WorkspaceUtils.DropDownList(GCD2[14]); 
//     }
////Client_Specific_Logo.Keys(GCD2[11]);
//
//
//
//var Job_Surcharge_Rule = Sys.Process("Maconomy").SWTObject("Shell", "Create Global Client").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 16).SWTObject("McValuePickerWidget", "", 2); 
//if(GCD2[15]!=""){
//  Job_Surcharge_Rule.Click();
//  WorkspaceUtils.SearchByValue(Job_Surcharge_Rule,"Job Surcharge Rule",GCD2[15]);
//         }

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



function todo(lvl,clientLvl){ 
    var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);

  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

  
  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();

  
  
  aqUtils.Delay(15000, Indicator.Text);
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
if(clientLvl==(ApproveInfo.length-1)){
if(lvl==3){
Client_Managt.ClickItem("|Approve Customer by Type (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Customer by Type (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Customer by Type (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Customer by Type (*)");
}
}
else{ 
 if(lvl==3){
Client_Managt.ClickItem("|Approve Customer (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Customer (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Customer (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Customer (*)");
} 
}
break;
}
}
}
}


function CredentialLogin(){ 
for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 

     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }
  else{ 
   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
    if(UserN){ 
      goToHR();
      UserN = false;
    }
    temp = searchNumber(Eno);
  }
//  Log.Message(temp)
  if(temp.length!=0){
    temp = temp+"*"+j;
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
  break;
  }
  }
  if((temp=="")||(temp==null))
  Log.Error("User Name is Not available for level :"+i);
//  Log.Message("Logins :"+temp);
}
WorkspaceUtils.closeAllWorkspaces();

}


function FinalApproveClient(comID,cltID,apvr,clientLvl){ 
//function FinalApproveClient(){
//  var ApproveInfo =[];
//  ApproveInfo[0] = "7";
//  ApproveInfo[1] = "7";
//  var comID ="1307";
//  var cltID = "111314";
//  var apvr ="CHFP CT Clients";
//  var clientLvl ="1";
  aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet.Show_Filter.Click();
}


var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.setText(cltID);
//firstCell.Keys("[Tab]");
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
  
//var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
//  job.setText(empNumber)
//job.setText(FullName + " "+STIME);
aqUtils.Delay(6000, Indicator.Text);;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==cltID){ 
//if(table.getItem(v).getText_2(2).OleValue.toString().trim()=="GAIL C COUTINHO"){

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}
    
    
    
    

ValidationUtils.verify(flag,true,"Created Client is available in system");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;
var Approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  Approve.Click();
  if(clientLvl==(ApproveInfo.length-1)){
 for(var j=0;j<12;j++){ 
    if(ImageRepository.ImageSet.Ok.Exists()){ 
//     ImageRepository.ImageSet.Ok.Click();
//     aqUtils.Delay(1000, Indicator.Text);;
var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    Ok.HoverMouse(); 
    ReportUtils.logStep_Screenshot();
    Ok.Click(); 
    aqUtils.Delay(8000, Indicator.Text); ;
   }
   else if(ImageRepository.ImageSet.OK_Button.Exists()){ 
//     ImageRepository.ImageSet.OK_Button.Click();
//     aqUtils.Delay(1000, Indicator.Text);;
var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    Ok.HoverMouse(); 
    ReportUtils.logStep_Screenshot();
    Ok.Click(); 
    aqUtils.Delay(8000, Indicator.Text); ;
   }
 }
    for(var j=0;j<12;j++){ 
if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Approve Customer by Type"){
var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    Ok.HoverMouse(); 
    ReportUtils.logStep_Screenshot();
    Ok.Click(); 
    aqUtils.Delay(8000, Indicator.Text); ;
}
}
      
}
 aqUtils.Delay(2000, Indicator.Text);;
if(ImageRepository.ImageSet.Maximize.Exists()){
ImageRepository.ImageSet.Maximize.Click();
}else{ 
var sideBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
sideBar.Click();
 aqUtils.Delay(2000, Indicator.Text);;
 ImageRepository.ImageSet.Maximize.Click();
}
var AllApproved = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.CTClientApproval;
AllApproved.Click();
aqUtils.Delay(8000, Indicator.Text); ;
ReportUtils.logStep_Screenshot();
var closeInfor = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
closeInfor.Click();
var showfilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
showfilter.Click();
aqUtils.Delay(5000, Indicator.Text); ;
var activeCustomer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Customers")
activeCustomer.Click();
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1)
.SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3)
.SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
firstCell.Click();
firstCell.setText("^a[BS]");
firstCell.setText(cltID);
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
aqUtils.Delay(6000, Indicator.Text);;
var flag=false;
if(table.getItemCount()==3){ 
  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==cltID){ 
  aqUtils.Delay(8000, Indicator.Text);;
  ReportUtils.logStep_Screenshot();
  ExcelUtils.setExcelName(workBook,"Data Management", true);
      ExcelUtils.WriteExcelSheet("Brand Number",EnvParams.Opco,"Data Management",table.getItem(1).getText_2(0).OleValue.toString().trim())
      ExcelUtils.WriteExcelSheet("Product Number",EnvParams.Opco,"Data Management",table.getItem(2).getText_2(0).OleValue.toString().trim())
  }

}
aqUtils.Delay(2000, Indicator.Text);;
  }
  }


}
