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
var sheetName = "CreateCurrencyJournal";
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
var CompanyNumber,Entrydate,CreatedOn1 = "";
//var JournalNumber = "";



function CreateCurrencyJournal(){
TextUtils.writeLog("Create a currency Journal Started"); 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//var Project_manager = ExcelUtils.getRowDatas("Agency - Senior Finance",EnvParams.Opco)
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
sheetName = "CreateCurrencyJournal";
Language = "";
ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
CompanyNumber,Entrydate = "";
Approve_Level = [];

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}

Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Copy GL  started::"+STIME);
TextUtils.writeLog("Execution Started :"+STIME);
getDetails();

//clientName = "1707_AutoClient 25February2020 10:41:17";
//brand = "AutoGlobalBrand 25February2020 10:41:17";
//product = "AutoGlobalProduct 25February2020 10:41:17" ;
//ClientNumber = "107743"

gotoMenu(); 
createcurrencyjournal();
//test();
//gl2();

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
sheetName="CreateCurrencyJournal";
  aqUtils.Delay(3000, Indicator.Text);
Entrydate = ExcelUtils.getRowDatas("DateEntry",EnvParams.Opco)
if((Entrydate==null)||(Entrydate=="")){ 
ValidationUtils.verify(false,true,"DateEntry is Needed to Create a Client");
}
Log.Message(Entrydate)

CompanyNumber = ExcelUtils.getRowDatas("CompanyNo",EnvParams.Opco)
Log.Message(CompanyNumber)
  if((CompanyNumber=="")||(CompanyNumber==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  CompanyNumber = ReadExcelSheet("Company No",EnvParams.Opco,"Data Management");
  Log.Message(CompanyNumber)
  }
    
CreatedOn1 = ExcelUtils.getRowDatas("CreatedOn",EnvParams.Opco)
Log.Message(CreatedOn1)
  if((CreatedOn1=="")||(CreatedOn1==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  CreatedOn1 = ReadExcelSheet("CreatedOn",EnvParams.Opco,"Data Management");
  Log.Message(CreatedOn1)
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

function createcurrencyjournal(){
//var gl=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
// waitForObj(gl);
//Sys.HighlightObject(gl);
//gl.Click();  

var currencyreval=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(currencyreval);
Sys.HighlightObject(currencyreval);
currencyreval.Click(); 

var printcurrencyreport=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
waitForObj(printcurrencyreport);
Sys.HighlightObject(printcurrencyreport);
printcurrencyreport.Click();


var company=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget;
Sys.HighlightObject(company);
company.Click();
company.setText(CompanyNumber);


//var all=Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
//
////Aliases.Maconomy.ReverseGL.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button
//waitForObj(all);
//Sys.HighlightObject(all);
//all.Click();

var journal=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McTextWidget2;
Sys.HighlightObject(journal);
journal.Click();
journal.setText(CompanyNumber);

var createglentries=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McPlainCheckboxView.Button;
Sys.HighlightObject(createglentries);
createglentries.Click();
//journal.setText(CompanyNumber);



if(Entrydate!=""){
    var date1=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite2.McDatePickerWidget;
Sys.HighlightObject(date1);
Log.Message(Entrydate);
WorkspaceUtils.CalenderDateSelection(date1,Entrydate)

      ValidationUtils.verify(true,true,"Date is selected in Maconomy"); 
    }
    else{ 
      ValidationUtils.verify(false,true,"Date is Needed to Create a Employee");
    } 

    var save=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
Sys.HighlightObject(save);
save.Click();
aqUtils.Delay(3000, Indicator.Text);

//var print=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//Sys.HighlightObject(print);
//print.Click();
    

var history=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
Sys.HighlightObject(history);
history.Click();

var journalnumber=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;

Sys.HighlightObject(journalnumber);
journalnumber.Click();
journalnumber.Keys("[Tab]");

var CompanyNumber1 = Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget; 
   waitForObj(CompanyNumber1);
  Sys.HighlightObject(CompanyNumber1)
  CompanyNumber1.Click(); 
CompanyNumber1.setText(CompanyNumber);
  //WorkspaceUtils.SearchByValue(CompanyNumber1,"Company",EnvParams.Opco,"Company Number");
CompanyNumber1.Keys("[Tab][Tab]");
  
var createdon=Aliases.Maconomy.Createcurrency.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget3;
  Sys.HighlightObject(createdon)
  createdon.Click();
  createdon.setText(CreatedOn1);
//var crgl=Aliases.Maconomy.Reverseok.CopyGL.Composite.Composite2.Composite.Button;
//crgl.Click();
//aqUtils.Delay(3000, Indicator.Text);

//var popup=Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal");
////Log.Message(popup.getText());
//
//var label=SWTObject("Label", "Journal no. * was created").getText();
//aqUtils.Delay(3000, Indicator.Text);
//log.Message(label);
//ExcelUtils.setExcelName(workBook,"Data Management", true);
//    ExcelUtils.WriteExcelSheet("ReserveJournalNo",EnvParams.Opco,"Data Management",label) 

}
