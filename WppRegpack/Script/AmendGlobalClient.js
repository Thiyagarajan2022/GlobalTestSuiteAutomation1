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
var sheetName = "AmendGlobalClient";
var ClientNo,BrandNo,Currency,Add1,Add2,Add3,Phone,Email ="";
var Approve_Level =[];
var ApproveInfo = [];
var Project_manager="";
var level =0;
function AmendGlobalClient(){ 
//  TextUtils.writeLog("Amend Gloabl brand Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
//var Project_manager = EnvParams.Opco+" Finance";
Project_manager = ExcelUtils.getRowDatas("Central Team - Client Account Management","Username")
Log.Message(Project_manager);
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
sheetName = "AmendGlobalClient";
ClientNo,BrandNo,Currency,Add1,Add2,Add3,Phone,Email ="";
Approve_Level =[];
ApproveInfo = [];
ExcelUtils.setExcelName(workBook, sheetName, true);
ClientNo = ExcelUtils.getRowDatas("Client Number",EnvParams.Opco)
  if((ClientNo=="")||(ClientNo==null)){
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  ClientNo = ReadExcelSheet("Client Number",EnvParams.Opco,"Data Management");
  }
if((ClientNo==null)||(ClientNo=="")){ 
ValidationUtils.verify(false,true,"Client Number is Needed to Amend Global Client");
}

Currency = ExcelUtils.getRowDatas("Currency",EnvParams.Opco)
if((Currency==null)||(Currency=="")){ 
ValidationUtils.verify(false,true,"Currency is Needed to Amend Global Client");
}

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Amend Global brand started::"+STIME);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
gotoMenu();
getDetails();
gotoClientSearch();
globalClient();
client();
WorkspaceUtils.closeAllWorkspaces();
CredentialLogin();
for(var i=level;i<ApproveInfo.length;i++){
level=i;
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);
var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);
todo(temp[3]);
FinalApprove(temp[1],temp[2]);
}
WorkspaceUtils.closeAllWorkspaces();
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
ReportUtils.logStep("INFO", "Moved to Client Management from Accounts Receivable Menu");
TextUtils.writeLog("Entering into Client Management from Accounts Receivable Menu");
}

function gotoClientSearch(){ 
 var CompanyNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.McValuePickerWidget;
  CompanyNumber.Click();
  WorkspaceUtils.SearchByValue(CompanyNumber,"Company",EnvParams.Opco,"Company Number");

 var curr = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.McPopupPickerWidget;
 curr.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
 if(Currency!=""){
  curr.Click();
  WorkspaceUtils.DropDownList(Currency,"Currency")
  }
  aqUtils.Delay(2000, Indicator.Text);
  
 var ClientNumber = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McValuePickerWidget;
  if(ClientNo!=""){
  ClientNumber.Click();
  WorkspaceUtils.VPWSearchByValue(ClientNumber,"Client",ClientNo,"Client Number");
    }
    
 var ClientName = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite4.McTextWidget;
 ClientName.setText("*");
 
 
 var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.RemarksSave;
 save.Click();
 aqUtils.Delay(5000, Indicator.Text);
 
 TextUtils.writeLog("Company Number, Client Number, Currency has entered and Saved in Client Search screen");
}

function globalClient(){ 
  var GblClient = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl2;
  GblClient.Click();
  aqUtils.Delay(3000, Indicator.Text);
  var active = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
  active.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  
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
  
  aqUtils.Delay(5000, Indicator.Text);
  TextUtils.writeLog("Global Client is available in maconomy to Amend");
}

function client(){ 
  var home = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.POApproval;
  home.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var information = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.TabControl;
  information.Click();
  aqUtils.Delay(2000, Indicator.Text);
  var C_Add1 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite5.McTextWidget;
  var C_Add2 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite6.McTextWidget;
  var C_Add3 = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite7.McTextWidget;
  var C_Phone = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite8.McTextWidget;
  var C_mail = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite9.McTextWidget;
  
  var changes = false;
  if(Add1!=""){
    if(C_Add1.getText()!=Add1){
    C_Add1.setText(Add1);
    ValidationUtils.verify(true,true,"Address 1 is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Address 1 in datasheet is as same as Value in Maconomy")
    }
    
  if(Add2!=""){
    if(C_Add2.getText()!=Add2){
    C_Add2.setText(Add2);
    ValidationUtils.verify(true,true,"Address 2 is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Address 2 in datasheet is as same as Value in Maconomy")
    }
    
    if(Add3!=""){
    if(C_Add3.getText()!=Add3){
    C_Add3.setText(Add3);
    ValidationUtils.verify(true,true,"Address 3 is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Address 3 in datasheet is as same as Value in Maconomy")
    }
    
    
    if(Phone!=""){
    if(C_Phone.getText()!=Phone){
    C_Phone.setText(Phone);
    ValidationUtils.verify(true,true,"Phone Number is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Phone Number in datasheet is as same as Value in Maconomy")
    }
    

    if(Email!=""){
    if(C_mail.getText()!=Email){
    C_mail.setText(Email);
    ValidationUtils.verify(true,true,"Email is Changed");
    changes = true;
    }
    else
    ReportUtils.logStep("INFO","Given Email in datasheet is as same as Value in Maconomy")
    }
    
    
    if(changes){ 
      var save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      save.Click();
      aqUtils.Delay(5000, "Waiting to save changes");
      var Submit = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.CloseFilter.Composite2.SingleToolItemControl;
      Submit.Click();
      aqUtils.Delay(5000, "Waiting to Submit changes");
      if(ImageRepository.ImageSet.Forward.Exists()){ 
      if(ImageRepository.ImageSet.Maximize.Exists()){
      ImageRepository.ImageSet.Maximize.Click();
        }
      }
      else{ 
        var ClientApproveBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.Approvals;
        ClientApproveBar.Click();
        ImageRepository.ImageSet.Maximize.Click();
      }
      var ClientApproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl3;
      ClientApproval.Click();
//      if(ClientApproval.getText()=="Client Approval"){ 
      var ApproverTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      Sys.HighlightObject(ApproverTable);
         var y=0;
     for(var i=0;i<ApproverTable.getItemCount();i++){   
     var approvers="";

      approvers = EnvParams.Opco+"*"+ClientNo+"*"+ApproverTable.getItem(i).getText_2(7).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(8).OleValue.toString().trim();
      Log.Message("Approver level :" +i+ ": " +approvers);
      ReportUtils.logStep("INFO","Approver Level "+i+": "+approvers);
//      Approve_Level[y] = "1307*1307100030*1307 SeniorFinance (13079510)*1307 Management (13079507)*"
      Approve_Level[y] = approvers;
      Log.Message(Approve_Level[y])
      y++;

    }
    var CloseBar = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
    CloseBar.Click();
    ImageRepository.ImageSet.Forward.Click();
    var OpCo1 = EnvParams.Opco;
    Log.Message(Approve_Level[0])
    var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
    if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
    level = 1;
    var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite7.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl5;
    Sys.HighlightObject(Approve)
    if(Approve.isEnabled()){ 
    Approve.HoverMouse();
    ReportUtils.logStep_Screenshot();
      Approve.Click();
    //ImageRepository.ImageSet.ApprovePurchaseOrder.Click();
    ReportUtils.logStep_Screenshot();
    aqUtils.Delay(8000, Indicator.Text);;
    ValidationUtils.verify(true,true,"Amend Global Client is Approved by "+Project_manager)
    TextUtils.writeLog("Levels 0 has  Approved Amend Global Client");
    aqUtils.Delay(8000, Indicator.Text);;
    }
    }
    
//      }
      
    }else{ 
      ValidationUtils.verify(false,true,"There is no changes happen in Maconomy screen")
    }
    
}


function getDetails(){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
Add1 = ExcelUtils.getRowDatas("Address_1",EnvParams.Opco)
Add2 = ExcelUtils.getRowDatas("Address_2",EnvParams.Opco)
Add3 = ExcelUtils.getRowDatas("Address_3",EnvParams.Opco)
Phone = ExcelUtils.getRowDatas("Phone",EnvParams.Opco)
Email = ExcelUtils.getRowDatas("E-mail",EnvParams.Opco)
if(((Add1==null)||(Add1==""))&&((Add2==null)||(Add2==""))&&((Add3==null)||(Add3==""))&&((Phone==null)||(Phone==""))&&((Email==null)||(Email==""))){ 
ValidationUtils.verify(false,true,"Any value from Address 1,Address 2,Address 3,Phone,Email is Needed to Amend Global Client");
}

}

function DropDownList(value){ 
var checkMark = false;
Sys.Process("Maconomy").Refresh();
  var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  var Add_Visible4 = true;
  while(Add_Visible4){
  if(list.isEnabled()){
  Add_Visible4 = false;
      for(var i=0;i<list.getItemCount();i++){ 
        if(list.getItem(i).getText_2(0)!=null){ 
          if(list.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
            list.Keys("[Enter]");
            aqUtils.Delay(5000, Indicator.Text);;
            checkMark = true;
            ValidationUtils.verify(true,true,"Yes is selected in Blocked Status");
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
//  else{ 
//   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
//    if(UserN){ 
//      goToHR();
//      UserN = false;
//    }
//    temp = searchNumber(Eno);
//  }
//  Log.Message(temp)
  if(temp.length!=0){
    temp = temp+"*"+j;
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
  break;
  }
  }
  if((temp=="")||(temp==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message("Logins :"+temp);
}
WorkspaceUtils.closeAllWorkspaces();

}




function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo = Aliases.Maconomy.Shell.Composite.Composite.Composite.TodoGrid.PTabFolder.TabFolderPanel.ToDo;
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

try{
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
catch(e){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
try{
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
catch(e){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}


if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Customer by Type (")!=-1)&&(temp1.length==2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into Approve Customer by Type from To-Dos List"); 
  }
}
if(lvl==3)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf("Approve Customer by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Customer by Type (Substitute) from To-Dos List");   
  }
}  

}


function FinalApprove(C_Num,Apvr){ 
  aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet.Show_Filter.Click();
}

var table = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid;
var firstCell = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.McValuePickerWidget;
firstCell.setText(C_Num);
var closefilter = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.TabFolderPanel.Composite.CloseFilter;
  
aqUtils.Delay(6000, Indicator.Text);;
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(0).OleValue.toString().trim()==C_Num){ 
    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

ValidationUtils.verify(flag,true,"Amended Client is available in Approval List");
TextUtils.writeLog("Amended Client is available in Approval List");
if(flag){ 
closefilter.HoverMouse();
ReportUtils.logStep_Screenshot();
closefilter.Click();
aqUtils.Delay(5000, Indicator.Text);;

var Approve = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl4;
Sys.HighlightObject(Approve)
if(Approve.isEnabled()){ 
Approve.HoverMouse();
ReportUtils.logStep_Screenshot();
Approve.Click();
aqUtils.Delay(9000, "Waiting to Approve");;
ValidationUtils.verify(true,true,"Amended Client is Approved by "+Apvr)
TextUtils.writeLog("Amended Client is Approved by "+Apvr);
}
}
}
