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
var sheetName = "ReverseGL";
var level =0;
var Approve_Level = []; 
var ApproveInfo = [];
var mainParent = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var newjournal =[];
var JournalNumber,Entrydate = "";
//var JournalNumber = "";



function ReverseGL(){
TextUtils.writeLog("Reverse GL Started"); 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

ExcelUtils.setExcelName(workBook, "Agency Users", true);
var Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
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
sheetName = "ReverseGL";
Language = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
JournalNumber,Entrydate = "";
Approve_Level = [];

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}

Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Reverse GL  started::"+STIME);
TextUtils.writeLog("Execution Started :"+STIME);
getDetails();

//clientName = "1707_AutoClient 25February2020 10:41:17";
//brand = "AutoGlobalBrand 25February2020 10:41:17";
//product = "AutoGlobalProduct 25February2020 10:41:17" ;
//ClientNumber = "107743"

gotoMenu(); 
gltransaction();
test();
gl2();

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
//WorkspaceUtils.closeAllWorkspaces();
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
if(ImageRepository.ImageSet.GendralLedger.Exists()){
ImageRepository.ImageSet.GendralLedger.Click();// GL
}
else if(ImageRepository.ImageSet.GendralLedger1.Exists()){
ImageRepository.ImageSet.GendralLedger1.Click();
}
else{
ImageRepository.ImageSet.GendralLedger2.Click();
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
Client_Managt.ClickItem("|GL Transactions");
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|GL Transactions");
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to GL Transaction from General Ledger Menu");
TextUtils.writeLog("Entering into GL Transaction from General Ledger Menu");
}


function getDetails(){ 
Indicator.PushText("Reading Data from Excel");
ExcelUtils.setExcelName(workBook, sheetName, true);
sheetName="ReverseGL";
  aqUtils.Delay(3000, Indicator.Text);
Entrydate = ExcelUtils.getRowDatas("DateEntry",EnvParams.Opco)
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"DateEntry is Needed to Create a Client");
}
Log.Message(Entrydate)

JournalNumber = ExcelUtils.getRowDatas("JournalNo",EnvParams.Opco)
Log.Message(JournalNumber)
  if((JournalNumber=="")||(JournalNumber==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  JournalNumber = ReadExcelSheet("General Journal No",EnvParams.Opco,"Data Management");
  Log.Message(JournalNumber)
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

function gltransaction(){
var gl1=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(gl1);
Sys.HighlightObject(gl1);
gl1.Click(); 

var gl=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(gl);
Sys.HighlightObject(gl);
gl.Click();

var all=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;

//Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button
waitForObj(all);
Sys.HighlightObject(all);
all.Click();

var journal=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(journal);
journal.Click();
journal.setText(JournalNumber);

var closefilter=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(closefilter);
closefilter.Click();
aqUtils.Delay(3000, Indicator.Text);

var cgl=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(cgl);
cgl.Click();

if(Entrydate!=""){
    var date1=Aliases.Maconomy.ReverseGL1.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.McGroupWidget.Composite.McDatePickerWidget;
Sys.HighlightObject(date1);
Log.Message(Entrydate);
WorkspaceUtils.CalenderDateSelection(date1,Entrydate)

      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Create a Employee");
    } 


var crgl=Aliases.Maconomy.ReverseGL1.Composite.Composite.Composite2.Composite.Button;
crgl.Click();
aqUtils.Delay(3000, Indicator.Text);

//var popup=Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal");
////Log.Message(popup.getText());
//
//var label=SWTObject("Label", "Journal no. * was created").getText();
//aqUtils.Delay(3000, Indicator.Text);
//log.Message(label);
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//    ExcelUtils.WriteExcelSheet("ReserveJournalNo",EnvParams.Opco,"Data Management",label) 

}


function test(){
var popup=Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal");
Sys.HighlightObject(popup);
//Log.Message(popup.getText());

var label=Aliases.Maconomy.SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Label", "*");
aqUtils.Delay(3000, Indicator.Text);
Sys.HighlightObject(label);
Log.Message(label.getText());
var gg = label.getText().OleValue.toString().trim();
var newjournal = gg.substring((gg.indexOf("Journal no. ")+12),gg.indexOf(" was created"))
Log.Message(newjournal);
//var newjournal = popmessage.split(" ");
//
//Log.Message(newjournal[0])
//Log.Message(newjournal[1])
//Log.Message(newjournal[2])
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("ReserveJournalNo",EnvParams.Opco,"Data Management",newjournal) 
  
var ok= Aliases.Maconomy.ReverseGLOK.Composite.Button;

}

function gl2(){
  
var generaljournal1= Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl2;
Sys.HighlightObject(generaljournal1);
generaljournal1.Click();

var companyno=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(companyno);
companyno.Click();
companyno.Keys("[Tab]");

aqUtils.Delay(5000, Indicator.Text);
ExcelUtils.setExcelName(workBook, "Data Management", true);
  JournalNumber1 = ReadExcelSheet("ReserveJournalNo",EnvParams.Opco,"Data Management");
  Log.Message(JournalNumber1)
var glno= Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget2;
waitForObj(glno);
Sys.HighlightObject(glno);
glno.Click();
glno.setText(JournalNumber1);
aqUtils.Delay(5000, Indicator.Text);

var closefilter1=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(closefilter1);
closefilter1.Click();

var email=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(email);
email.Click();


var pop= Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal")
Sys.HighlightObject(pop);

var ok= Aliases.Maconomy.ReverseGLOK.Composite.Button;
ok.Click();
}


//function test1()
//
//{
//  
////SWTObject("Shell", "GL Transactions - General Journal")
//   if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="GL Transactions - General Journal")    
//    {
//   //   Aliases.Maconomy.SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Label", "Journal no. 1707117407 was created")
//  //  var button = Aliases.Maconomy.CreateGlobalBrand3.Composite.Button;
//    //Aliases.CompanyRegistrationAlreadyUsED
//   //   var label =Aliases.Maconomy.CreateGlobalBrand3.SWTObject("Label", "*").WndCaption;
//   
//   var label = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal",1).SWTObject("Label", "*");
//  
//
//   //  NameMapping.Sys.Maconomy.Shell.SWTObject("Label", "*").WndCaption;
//      Log.Message(label );
//   //    button.HoverMouse();
//    // ReportUtils.logStep_Screenshot("");
//  //    button.Click();
//      Delay(5000);
//  }
//
//}



//function gotoClientSearch(){ 
// var CompanyNumber = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
// waitForObj(CompanyNumber);
// 
//  Sys.HighlightObject(CompanyNumber);
//  CompanyNumber.Click();
//  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
//  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");
//
// var curr = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
// curr.Keys(" ");
// curr.HoverMouse();
// Sys.HighlightObject(curr);
// if(currency!=""){
//  curr.Click();
//  WorkspaceUtils.DropDownList(currency,"Currency")
//  }
////  aqUtils.Delay(2000, Indicator.Text);
//var ClientNumber = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McValuePickerWidget;
//  if(ClientNo!=""){
//  ClientNumber.Click();
//  WorkspaceUtils.VPWSearchByValue(ClientNumber,"Client",ClientNo,"Client Number");
//    }
//
// var ClientName = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;
// ClientName.HoverMouse();
// Sys.HighlightObject(ClientName);
// ClientName.setText("*");
// //ClientName.setText(clientName+" "+STIME);
// 
// 
// var save = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
// save.HoverMouse();
// Sys.HighlightObject(save);
// save.Click();
//// aqUtils.Delay(5000, Indicator.Text);
// 
// TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
//}
//
//
////function globalClient(){ 
//  var GblClient = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
//  GblClient.HoverMouse();
//  Sys.HighlightObject(GblClient);
//  GblClient.Click();
////  aqUtils.Delay(5000, Indicator.Text);
//  var AllClients = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
//  AllClients.Click();
//  AllClients.HoverMouse();
//  AllClients.HoverMouse();
//  AllClients.HoverMouse();
//  
//  aqUtils.Delay(3000, "Reading from Global Client table");
//  var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//  
//  if(table.getItem(0).getText_2(0).OleValue.toString().trim()==ClientNo){
//  table.HoverMouse(49, 52);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 52);
//  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
//  }
//  else if(table.getItem(1).getText_2(0).OleValue.toString().trim()==ClientNo){
//  table.HoverMouse(49, 71);
//  ReportUtils.logStep_Screenshot();  
//  table.Click(49, 71);
//  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
//  }
//  else if(table.getItem(2).getText_2(0).OleValue.toString().trim()==ClientNo){
//  table.HoverMouse(49, 90);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 90);
//  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
//  }
//  else if(table.getItem(3).getText_2(0).OleValue.toString().trim()==ClientNo){
//  table.HoverMouse(49, 109);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 109);
//  ValidationUtils.verify(true,true,"Global Client is available in maconomy to Amend");
//  }
//  
//  aqUtils.Delay(5000, "Playback");
//  TextUtils.writeLog("Global Client is available in maconomy to Amend");
//  
//  
//  }
//
//
//
////function newGlobalClient(){ 
//  
//
//var home= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
//// var ClientName = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.ClientName_textbox;
// waitForObj(home);
//
// home.Click();
// Sys.HighlightObject(home);
//// ClientName.setText(clientName.toString().trim()+" "+STIME);
//// clientName = clientName.toString().trim()+" "+STIME;
// 
//var sublevel= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
//// var ClientName = Aliases.Maconomy.New_Global_Client.Composite.Composite.Composite.Composite.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.ClientName_textbox;
// waitForObj(sublevel);
//
// sublevel.Click();
// Sys.HighlightObject(sublevel);
// 
// 
// var glbclient= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
//  waitForObj(glbclient);
//
// glbclient.Click();
// Sys.HighlightObject(glbclient);
// 
//  var newglbbrand= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//  waitForObj(newglbbrand);
//  newglbbrand.Click();
// Sys.HighlightObject(newglbbrand);
//
// 
//var brandname1=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
// // waitForObj(brandname1);
//  brandname1.Click();
//Sys.HighlightObject(brandname1);
//brandname1.setText(brandname);
//
//var defaultname1=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McTextWidget;
// // waitForObj(brandname1);
//  defaultname1.Click();
//Sys.HighlightObject(defaultname1);
//defaultname1.setText(defaultname);
//
////Default Name
// 
//var next=Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.Composite.Button;
// next.HoverMouse();
// Sys.HighlightObject(next);
// ReportUtils.logStep_Screenshot() ;
// next.Click();
// 
//
//}
//
////function GlobalClient_Screen2(){ 
//  var CompanyNumber = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite.McValuePickerWidget;
// waitForObj(CompanyNumber);
// 
//  Sys.HighlightObject(CompanyNumber);
//  CompanyNumber.Click();
//  var ExlArray = getExcelData_Company("Validate_Company",EnvParams.Opco)
//  WorkspaceUtils.config_with_Maconomy_Validation(CompanyNumber,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,ExlArray,"Company Number");
//
//
//  var C_Language = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite2.McPopupPickerWidget;
//  if(clientlan!=""){
//  C_Language.Click();
//  WorkspaceUtils.DropDownList(clientlan,"Language")
//  }
////  aqUtils.Delay(2000, Indicator.Text);
//  
//  var Attn = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite3.McValuePickerWidget;
//  Attn.HoverMouse();
//  Sys.HighlightObject(Attn);
//  Attn.setText(attn);  
//  
//  var C_Email  = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite4.McTextWidget;
////  var Eml_split1 = mail.OleValue.toString().trim().substring(0,mail.OleValue.toString().trim().indexOf("@"));
////  var Eml_split2 = mail.OleValue.toString().trim().substring(mail.OleValue.toString().trim().indexOf("@"));
////  Eml_split1 = Eml_split1 +" "+STIME;
////  Eml_split1 = Eml_split1.replace(/[_: ]/g,"");
////  mail = Eml_split1+Eml_split2
//  C_Email.setText(mail);
//  
//  var C_phone = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite5.McTextWidget;
//  C_phone.setText(phone); 
//  
//  var C_AcctDir = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite6.McValuePickerWidget;
//  if(AccDir!=""){
//  C_AcctDir.HoverMouse();
//  Sys.HighlightObject(C_AcctDir);
//  C_AcctDir.Click();
//  WorkspaceUtils.SearchByValue(C_AcctDir,"Employee",AccDir,"Acct Director No");
//  }
//  
//  var C_PaymentTerm = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite7.McPopupPickerWidget;
//  if(payterm!=""){
//  Sys.HighlightObject(C_PaymentTerm);
//  C_PaymentTerm.Click();
//  WorkspaceUtils.DropDownList(payterm,"Payment Terms")
//  }
////  aqUtils.Delay(2000, Indicator.Text);
//  
//  var C_companyTaxCode = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite8.McPopupPickerWidget;
//  if(Comtaxcode!=""){
//  C_companyTaxCode.HoverMouse();
//  Sys.HighlightObject(C_companyTaxCode);
//  C_companyTaxCode.Click();
//  WorkspaceUtils.DropDownList(Comtaxcode,"Company Tax Code");
//  }
////  aqUtils.Delay(2000, Indicator.Text);
//  
//  var C_JobPriceList = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite.Composite.Composite2.Composite.McPaneGui_10.Composite.Composite.McGroupWidget2.Composite9.McValuePickerWidget;
//  if(sales!=""){
//  Sys.HighlightObject(C_JobPriceList);
//  C_JobPriceList.Click();
//  WorkspaceUtils.SearchByValue(C_JobPriceList,"Job Price List",sales,"Job Price List Sales");
//         }  
//         
////    aqUtils.Delay(2000, Indicator.Text);
//  
// var Next = Aliases.Maconomy.CreateGlobalBrand2.Composite.Composite.Composite2.Composite.Button;
// Sys.HighlightObject(Next);
// Next.HoverMouse();
// ReportUtils.logStep_Screenshot() ;
// Next.Click();
//    Delay(5000);
//
//}
//
////
////function test()
//{
//  
// if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Global Client - Client Information Card")    
//    {
//    var button = Aliases.Maconomy.CreateGlobalBrand3.Composite.Button;
//    //Aliases.CompanyRegistrationAlreadyUsED
//      var label =Aliases.Maconomy.CreateGlobalBrand3.SWTObject("Label", "*").WndCaption;
//     //  NameMapping.Sys.Maconomy.Shell.SWTObject("Label", "*").WndCaption;
//      Log.Message(label );
//       button.HoverMouse();
//    // ReportUtils.logStep_Screenshot("");
//      button.Click();
//      Delay(5000);
//  }
////      
//    
// if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Global Client - Client Information Card")    
//    { 
//  var button1=Aliases.Maconomy.CreateGlobalBrand3.Composite.Button;
//  var label =Aliases.Maconomy.CreateGlobalBrand3.SWTObject("Label", "*").WndCaption;
//  Log.Message(label );
//       button1.HoverMouse();
//    // ReportUtils.logStep_Screenshot("");
//      button1.Click();
//      Delay(5000);
//}
//}
//
//
//
////function brand1(){
////var name1=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
////name1.Click();
////  name1.Keys(brandname);  
//var name1=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
//name1.Click();
//  name1.Keys(brandname); 
//var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
// Sys.HighlightObject(table);
////Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid
// 
//  aqUtils.Delay(4000, Indicator.Text);
//    
//  if(table.getItem(0).getText_2(1).OleValue.toString().trim()==brandname){
//  table.HoverMouse(49, 52);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 52);
//  ValidationUtils.verify(true,true,"Global Brand is available in maconomy to block");
//  }
//  else if(table.getItem(1).getText_2(1).OleValue.toString().trim()==brandname){
//  table.HoverMouse(49, 71);
//  ReportUtils.logStep_Screenshot();  
//  table.Click(49, 71);
//  ValidationUtils.verify(true,true,"Global Brand is available in maconomy to block");
//  }
//  else if(table.getItem(2).getText_2(1).OleValue.toString().trim()==brandname){
//  table.HoverMouse(49, 90);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 90);
//  ValidationUtils.verify(true,true,"Global Brand is available in maconomy to block");
//  }
//  else if(table.getItem(3).getText_2(1).OleValue.toString().trim()==brandname){
//  table.HoverMouse(49, 109);
//  ReportUtils.logStep_Screenshot();
//  table.Click(49, 109);
//  ValidationUtils.verify(true,true,"Global Brand is available in maconomy to block");
//  }
//  
//  aqUtils.Delay(5000, Indicator.Text);
//
//  TextUtils.writeLog("Global Brand is available in maconomy to block");
//  
//  
//}
//
//
//  
////function indiaSpecific(){ 
//  aqUtils.Delay(7000, Indicator.Text);
//  var indiaspec = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.IndiaSpecific;
//Sys.HighlightObject(indiaspec);
//var Start = StartwaitTime();
//var waitTime = true;
//var Difference = 0;
//while(waitTime)
//if(Difference<61){
//if((indiaspec.isEnabled())&&(indiaspec.text=="India Specific")){
//Sys.HighlightObject(indiaspec);
//indiaspec.HoverMouse();
//indiaspec.Click();
//waitTime = false;
//}
//else{ 
//var End = EndTime();
//Difference = End - Start;
//}
//}
//else{
// ValidationUtils.verify(true,false,"Screen is not Responding more than a minute");
//}
//
//  var StateCode = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McPopupPickerWidget;
//  var debtorType = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
//  var C_pan = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.PAN;
//  var C_tan = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.TAN;
//    
//  if(State!=""){
//  Sys.HighlightObject(StateCode);
//  StateCode.HoverMouse();
//  StateCode.Click();
//  DropDownList(State.OleValue.toString().trim(),"State Code")
//  }
//  aqUtils.Delay(2000, Indicator.Text);
//  
//  if(GST!=""){
//  Sys.HighlightObject(debtorType);
//  debtorType.HoverMouse();
//  debtorType.Click();
//  WorkspaceUtils.DropDownList(GST,"GST Debtor Type")
//  }
//  
//  if(PAN!=""){
//  Sys.HighlightObject(C_pan);
//  C_pan.HoverMouse();  
//   C_pan.setText(PAN.OleValue.toString().trim());
//  }
//  
//  if(TAN!=""){
//   C_tan.setText(TAN.OleValue.toString().trim());
//  }
//var save = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.Save;
//Sys.HighlightObject(save);
//save.HoverMouse();
//save.Click();
//
//}
//  
// 
//  
////function Information(){ 
//  
//  var info = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl2;
//  info.HoverMouse();
//  info.HoverMouse();
//  info.HoverMouse();
//  Sys.HighlightObject(info);
//  info.HoverMouse();
//  info.HoverMouse();
//  info.Click();
//  aqUtils.Delay(2000, Indicator.Text);
//  var submit = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//  Sys.HighlightObject(submit);
//  submit.HoverMouse();
//  submit.HoverMouse();
//  submit.Click();
//}
//  
////function ApprvalInformation(){ 
// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(ClientApproval);
// ClientApproval.HoverMouse();
// ClientApproval.Click();
// if(ImageRepository.ImageSet.Maximize.Exists()){
//ImageRepository.ImageSet.Maximize.Click();
//}
// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(ClientApproval);
// ClientApproval.HoverMouse();
// ClientApproval.Click();
//   var ApproverTable = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//   var y=0;
//  for(var i=0;i<ApproverTable.getItemCount();i++){   
//     var approvers="";
//      if(ApproverTable.getItem(i).getText_2(3)!="Approved"){
//      approvers = EnvParams.Opco+"*"+client1+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(5).OleValue.toString().trim();
//      Log.Message("Approver level :" +i+ ": " +approvers);
////      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
//      Approve_Level[y] = approvers;
//      y++;
//      }
//}
//
//TextUtils.writeLog("Finding approvers for Created Global Client");
//var closeCAList = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
// Sys.HighlightObject(closeCAList);
// closeCAList.HoverMouse();
// closeCAList.Click();
// 
//ImageRepository.ImageSet.Forward.Click();
//
//
//CredentialLogin();
//var OpCo2 = ApproveInfo[0].split("*");
////var OpCo1 = EnvParams.Opco;
//ExcelUtils.setExcelName(workBook, "Server Details", true);
//var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
////var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
////var sheetName = "Agency Users";
////workBook = Project.Path+excelName;
////ExcelUtils.setExcelName(workBook, sheetName, true);
////OpCo2 = ExcelUtils.AgencyLogin(OpCo2,EnvParams.Opco);
//sheetName = "CreateGlobalBrand";
//if(OpCo2[2]==Project_manager){
//level = 1;
//var Approve = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.SingleToolItemControl;
//Sys.HighlightObject(Approve);
//if(Approve.isEnabled()){ 
//Approve.HoverMouse();
//ReportUtils.logStep_Screenshot();
//Approve.Click();
//aqUtils.Delay(8000, "Waiting for Approve");;
//ValidationUtils.verify(true,true,"Purchase Order is Approved by "+Project_manager)
//TextUtils.writeLog("Levels 0 has  Approved the Created Budget");
////aqUtils.Delay(8000, Indicator.Text);;
//}
//}
////var Approve = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.Approve;
//// Sys.HighlightObject(Approve);
//// Approve.HoverMouse();
//// Approve.Click();
//}
//  
//  
////function CredentialLogin(){ 
//  var AppvLevl = [];
//for(var i=0;i<Approve_Level.length;i++){
//  var UserN = true;
//  var temp="";
//  var temp1="";
//  var Cred = Approve_Level[i].split("*");
//  for(var j=2;j<4;j++){
//  temp="";
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
//  { 
//     var sheetName = "Agency Users";
//     workBook = Project.Path+excelName;
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
//  }
//  else if((Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
//  { 
//
//    var sheetName = "SSC Users";
//    ExcelUtils.setExcelName(workBook, sheetName, true);
//    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
//  }
//
//  if(temp.length!=0){
//    temp1 = temp1+temp+"*"+j+"*";
////  break;
//  }
//  }
//  if((temp1=="")||(temp1==null))
//  Log.Error("User Name is Not available for level :"+i);
//  Log.Message(temp1)
//  AppvLevl[i] = temp1;
//}
//  ApproveInfo = levelMatch(AppvLevl)
//  Log.Message("-----Approvers-------------")
//  for(var i=0;i<ApproveInfo.length;i++){
//    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
//    Log.Message(ApproveInfo[i]);
//    }
////WorkspaceUtils.closeAllWorkspaces();
//}
//
//  
////function todo(lvl){ 
//  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
//  var toDo = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
//  toDo.HoverMouse();
//  ReportUtils.logStep_Screenshot();
//  toDo.DBlClick();
//  TextUtils.writeLog("Entering into To-Dos List");
//  aqUtils.Delay(3000, Indicator.Text);
//  //To Maximaize the window
//  Sys.Desktop.KeyDown(0x12);
//  Sys.Desktop.KeyDown(0x20);
//  Sys.Desktop.KeyUp(0x12);
//  Sys.Desktop.KeyUp(0x20);
//  Sys.Desktop.KeyDown(0x58);
//  Sys.Desktop.KeyUp(0x58);  
//  aqUtils.Delay(1000, Indicator.Text);
//
//var refresh= Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//  
////if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
////var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
////}
////if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
////var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
////}
//refresh.Click();
//aqUtils.Delay(15000, Indicator.Text);
//var Client_Managt=Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
//
////if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
////Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
////}
////if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
////Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
////}
//var listPass = true;
//if(lvl==2)
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Customer by Type (")!=-1)&&(temp1.length==2)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);  
//TextUtils.writeLog("Entering into Approve Customer by Type from To-Dos List");
//listPass = false; 
//  }
//}
//if(lvl==3)
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Customer by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
//Client_Managt.ClickItem("|"+temp);    
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp); 
//TextUtils.writeLog("Entering into Approve Customer by Type (Substitute) from To-Dos List");
//var listPass = true;   
//  }
//}  
//if(listPass){
//if(lvl==2)
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Customer (")!=-1)&&(temp1.length==2)){ 
//Client_Managt.ClickItem("|"+temp);   
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp);  
//TextUtils.writeLog("Entering into Approve Customer from To-Dos List");
//listPass = false; 
//  }
//}
//if(lvl==3)
//for(var j=0;j<Client_Managt.getItemCount();j++){ 
//  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
//  var temp1 = temp.split("(");
//if((temp.indexOf("Approve Customer (Substitute) (")!=-1)&&(temp1.length==3)){ 
//Client_Managt.ClickItem("|"+temp);    
//ReportUtils.logStep_Screenshot(); 
//Client_Managt.DblClickItem("|"+temp); 
//TextUtils.writeLog("Entering into Approve Customer (Substitute) from To-Dos List");
//var listPass = true;   
//  }
//} 
//  }
//
//
//}
//
//  
////function FinalApproveClient(ClientNum,Apvr,lvl){ 
////aqUtils.Delay(5000, Indicator.Text);
////if(ImageRepository.ImageSet.Show_Filter.Exists()){
////aqUtils.Delay(2000, Indicator.Text);
////ImageRepository.ImageSet.Show_Filter.Click();
////}
//var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//waitForObj(table);
//Sys.HighlightObject(table);
//
//if(Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Visible){
//
//}else{ 
//var showFilter = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
//waitForObj(table);
//Sys.HighlightObject(showFilter);
//showFilter.HoverMouse();
//showFilter.HoverMouse();
//showFilter.HoverMouse();
//showFilter.Click();
//}
//
//var table = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
//var firstCell = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
//waitForObj(firstCell);
//Sys.HighlightObject(firstCell);
//firstCell.HoverMouse();
//firstCell.HoverMouse();
//firstCell.setText(ClientNum);
//aqUtils.Delay(3000, "Reading Data in table");;
//var closefilter = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
//waitForObj(closefilter);
//Sys.HighlightObject(closefilter);
//closefilter.HoverMouse();
//closefilter.HoverMouse(); 
//closefilter.HoverMouse();
//closefilter.HoverMouse(); 
////aqUtils.Delay(6000, Indicator.Text);;
//var flag=false;
//for(var v=0;v<table.getItemCount();v++){ 
//  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==ClientNum){ 
//    flag=true;    
//    break;
//  }
//  else{ 
//    table.Keys("[Down]");
//  }
//}
//
//ValidationUtils.verify(flag,true,"Created Client is available in Approval List");
//TextUtils.writeLog("Created Client is available in Approval List");
//if(flag){ 
//closefilter.HoverMouse();
//ReportUtils.logStep_Screenshot();
//closefilter.Click();
//aqUtils.Delay(5000, Indicator.Text);;
//
//var Approve = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//Sys.HighlightObject(Approve)
//if(Approve.isEnabled()){ 
//Approve.HoverMouse();
//ReportUtils.logStep_Screenshot();
//Approve.Click();
//aqUtils.Delay(8000, "Waiting To Approve");;
//ValidationUtils.verify(true,true,"Global Client is Approved by "+Apvr)
//aqUtils.Delay(8000, Indicator.Text);;
//TextUtils.writeLog("Global Client is Approved by "+Apvr);
//if(Approve_Level.length==lvl+1){
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//Ok.HoverMouse(); 
//ReportUtils.logStep_Screenshot();
//Ok.Click(); 
//aqUtils.Delay(8000, Indicator.Text); ;
// for(var j=0;j<12;j++){ 
//if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Approve Customer by Type"){ 
//var label = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Label", "*");
//Log.Message(label.getText());
//var lab = label.getText().OleValue.toString().trim();
//var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Approve Customer by Type").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//Ok.HoverMouse(); 
//ReportUtils.logStep_Screenshot();
//Ok.Click(); 
//aqUtils.Delay(8000, Indicator.Text); ;  
//}
// 
//
// }
// 
//  ExcelUtils.setExcelName(workBook,"Data Management", true);
//  ExcelUtils.WriteExcelSheet("Global Client",EnvParams.Opco,"Data Management",ClientNum)
//  TextUtils.writeLog("Global Client Number :"+ClientNum); 
//  
//// if(Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.Visible){
//// var ClientApproval = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.PTabItemPanel.ClientApproval;
// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(ClientApproval);
// ClientApproval.HoverMouse();
// ClientApproval.Click();
//// }
// if(ImageRepository.ImageSet.Maximize.Exists()){
//ImageRepository.ImageSet.Maximize.Click();
//}
// var ClientApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
// Sys.HighlightObject(ClientApproval);
// ClientApproval.HoverMouse();
// ClientApproval.Click();
//   var ApproverTable = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//  Sys.HighlightObject(ApproverTable);
//  ReportUtils.logStep_Screenshot();
//  var closeApproval = Aliases.Maconomy.CreateGlobalBrand1.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel.TabControl;
//  Sys.HighlightObject(closeApproval);
// closeApproval.HoverMouse();
// closeApproval.Click();
// ImageRepository.ImageSet.Forward.Click();
// var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//  menuBar.Click();
//}
//  ValidationUtils.verify(true,true,"Global Client is Approved by "+Apvr)
//
//  
//}
//}
//
//}  
//  
//
//function SearchByValue(ObjectAddrs,popupName,value,fieldName){ 
//var checkmark = false;
//  aqUtils.Delay(1000, popupName);;
//    Sys.Desktop.KeyDown(0x11);
//    Sys.Desktop.KeyDown(0x47);
//    Sys.Desktop.KeyUp(0x11);
//    Sys.Desktop.KeyUp(0x47);
//    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//  waitForObj(code);
//  code.Click();
//    code.setText(value);
////    aqUtils.Delay(3000, Indicator.Text);;
//    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Search").OleValue.toString().trim()+" ");
//    waitForObj(serch);
//
//  serch.Click();
////    aqUtils.Delay(5000, Indicator.Text);;
//  var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
//  var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
//
//
//    waitForObj(OK);
//    Sys.HighlightObject(table);
//    var itemCount = table.getItemCount();
//    if(itemCount>0){
//    for(var i=0;i<itemCount;i++){
//      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
//       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
//  waitForObj(OK);
//  OK.Click();
//
//          checkmark = true;
//          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
//          break;
//          
//      }
//      else{ 
//        Sys.Desktop.KeyDown(0x28);
//        Sys.Desktop.KeyUp(0x28);
//        if(i==itemCount-1){ 
//          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
//  waitForObj(cancel);
//  cancel.Click();
//
//          Sys.HighlightObject(ObjectAddrs);
//          ObjectAddrs.setText("");
//          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
//        }
//      }
//      
//      }
//    }
//    else { 
//      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
//        waitForObj(cancel);
//        cancel.Click();
//
//      Sys.HighlightObject(ObjectAddrs);
//      ObjectAddrs.setText("");
//      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
//    }
//    
//    return checkmark;
//}
//
//function VPWSearchByValue(ObjectAddrs,popupName,value,fieldName){ 
//var checkmark = false;
//  aqUtils.Delay(1000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x11);
//    Sys.Desktop.KeyDown(0x47);
//    Sys.Desktop.KeyUp(0x11);
//    Sys.Desktop.KeyUp(0x47);
////    aqUtils.Delay(3000, Indicator.Text);;
//
//    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
//  waitForObj(code);
//  code.Click();
//    code.setText(value);
////    aqUtils.Delay(3000, Indicator.Text);;
//    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Search").OleValue.toString().trim()+" ");
//    waitForObj(serch);
//
//  serch.Click();    
//    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
//    var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
//    waitForObj(OK);
//    Sys.HighlightObject(table); 
//    var itemCount = table.getItemCount();
//    if(itemCount>0){ 
//    for(var i=0;i<itemCount;i++){
//      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
//       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "OK").OleValue.toString().trim())
//  waitForObj(OK);
//  OK.Click();
//          checkmark = true;
//          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
//          break;
//          
//      }
//      else{ 
//        Sys.Desktop.KeyDown(0x28);
//        Sys.Desktop.KeyUp(0x28);
//        if(i==itemCount-1){ 
//          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
//  waitForObj(cancel);
//  cancel.Click();
//
//          Sys.HighlightObject(ObjectAddrs);
//          ObjectAddrs.setText("");
//          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
//        }
//      }
//      
//      }
//    }
//    else { 
//      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Cancel").OleValue.toString().trim());
//        waitForObj(cancel);
//        cancel.Click();
//
//      Sys.HighlightObject(ObjectAddrs);
//      ObjectAddrs.setText("");
//      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
//    }
//    return checkmark;
//}
//

