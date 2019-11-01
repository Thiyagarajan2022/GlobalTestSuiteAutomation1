//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "RejectingAbsenceRequest";
var AllowanceDetails = [];

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.TimeExpense.Exists()){
ImageRepository.ImageSet.TimeExpense.Click();// GL
}
else if(ImageRepository.ImageSet.TimeExpense1.Exists()){
ImageRepository.ImageSet.TimeExpense1.Click();
}
else{
ImageRepository.ImageSet.TimeExpense2.Click();
}

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|Absence");
}

}
Delay(6000);
//Allowance();
}


function Absence(){ 

  AllowanceDetails = SOXexcel(sheetName,1);

  var all = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Submitted");
  all.Click();
  Delay(4000);
  var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  if(AllowanceDetails[0]!=""){ 
    firstCell.Click();
    firstCell.setText(AllowanceDetails[0]);
  }
  firstCell.Keys("[Tab]");
  Delay(1000);
  var Duration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  if(AllowanceDetails[1]!=""){ 
    Duration.Click();
    Duration.setText(AllowanceDetails[1]);
  }
  Duration.Keys("[Tab][Tab]");
  Delay(1000);
  var lastDate = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 1);
  if(AllowanceDetails[2]!=""){ 
    lastDate.Click();
    lastDate.setText(AllowanceDetails[2]);
  }
  lastDate.Keys("[Tab]");
  Delay(1000);
  var AbsenceType = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 3);
  if(AllowanceDetails[3]!=""){ 
    AbsenceType.Click();
    WorkspaceUtils.SearchByValue(AbsenceType,"Absence Type",AllowanceDetails[3])
  }else{ 
    ValidationUtils.verify(false,true,"Absence Type is Needed to Reject Absence Request");
  }
  AbsenceType.Keys("[Tab]");
  Delay(5000);
  var temporary = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 4);
  if(AllowanceDetails[4]!=""){ 
    temporary.Click();
    WorkspaceUtils.DropDownList(temporary,AllowanceDetails[4]);
  }
  temporary.Keys("[Tab]");
  Delay(1000);
  var reason = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
  if(AllowanceDetails[5]!=""){ 
    reason.Click();
    reason.setText(AllowanceDetails[5]);
  }
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
 for(var i=0;i<table.getItemCount();i++){ 
   if((table.getItem(i).getText_2(0).OleValue.toString().trim()==AllowanceDetails[0]) && (table.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(AllowanceDetails[1])!=-1) && (table.getItem(i).getText_2(4).OleValue.toString().trim()==AllowanceDetails[3])) { 
//     Log.Message("Absence is Successfully Created");
     var createdby = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
     Sys.HighlightObject(createdby);
     createdby.Click(); 
     var creat =  createdby.getText();   
     var approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
     Sys.HighlightObject(approver);
     approver.Click();
     
     var approve = []; 
     approve[0] = creat+"*"+creat+"*"+approver.getText()+"*";
     Log.Message(approve[0])
     Delay(3000);
     WorkspaceUtils.closeAllWorkspaces();
     HRData = WorkspaceUtils.goToHR();
     LoginEmp = WorkspaceUtils.Credentiallogin(Project.Path+excelName, "userRoles");

     UserPasswd = WorkspaceUtils.Login_Match(approve,LoginEmp,HRData,"");

     RestMaconomy(UserPasswd);
     break;
     break;
 }
 else{ 
   table.Keys("[Tab]");
 }
 }
  
}


function RestMaconomy(UserPasswd){ 
//var UserPasswd = [];
//UserPasswd[0] = "1707200085*Test Sample - Automation4 23January2019 10:05:10***CORE@WPP123*1";;
//UserPasswd[0] = "1707200085*Test Sample - Automation4 23January2019 10:05:10*SSC IN -  Biller*CORE@WPP123*1";
//UserPasswd[0] = "1706*Automation Client 19December2018 19:53:38*SSC IN -  CT Clients*CORE@WPP123";
Log.Message(UserPasswd.length);
for(var j=0;j<UserPasswd.length;j++){

var temp = UserPasswd[j];
var temp_user = temp.split("*");
var uname = temp_user[2]; 
Log.Message(uname)
var pwd = temp_user[3];
Log.Message(pwd)
WorkspaceUtils.Rests(uname,pwd);
   var ApproveFlag = true; 
gotTODOs_Approve(temp_user[4]);

Delay(8000);
if(ImageRepository.ImageSet.Show_Filter.Exists()){ 
  ImageRepository.ImageSet.Show_Filter.Click();
  Delay(2000);
}
var allowanceApprove = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Absence for Approval");
allowanceApprove.Click();
Delay(3000);
////var invoiceAllocations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
////  invoiceAllocations.Click();
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  firstcell.Click();
//  firstcell.setText(temp_user[0]);
//  Delay(2000);
  firstcell.Keys("[Tab]");

  var Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  Job.Click();
  Delay(2000);
  Job.Keys("[Tab]");
  
  var Date = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  Date.Click();
  Delay(2000);
  if(AllowanceDetails[0]!="")
  Date.setText(AllowanceDetails[0]);
  Date.Keys("[Tab]");
  
  var Duration = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  Duration.Click();
  Delay(2000);
  if(AllowanceDetails[1]!="")
  Duration.setText(AllowanceDetails[1]);
  Duration.Keys("[Tab][Tab]");
  
  var LastDate = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  LastDate.Click();
  Delay(2000);
  if(AllowanceDetails[2]!="")
  LastDate.setText(AllowanceDetails[2]);
  LastDate.Keys("[Tab]");
  
  var AbsenceType = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 1);
  AbsenceType.Click();
  Delay(2000);
  if(AllowanceDetails[3]!="")
  AbsenceType.setText(AllowanceDetails[3]);
  else{ 
    ValidationUtils.verify(false,true,"Absence Type is Needed to Reject Absence Request");
  }  
  AbsenceType.Keys("[Tab]");
  Delay(5000);
  var temporary = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 4);
  if(AllowanceDetails[4]!=""){ 
    temporary.Click();
    WorkspaceUtils.DropDownList(temporary,AllowanceDetails[4]);
  }
  temporary.Keys("[Tab]");
  Delay(1000);
  
  var Reason = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  Reason.Click();
  Delay(2000);
  if(AllowanceDetails[5]!="")
  Reason.setText(AllowanceDetails[5]);
  Reason.Keys("[Tab]");
  

  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
      var itemCount = table.getItemCount();
      var flag = false;
      for(var i=0;i<itemCount;i++){
      if((table.getItem(i).getText_2(2).OleValue.toString().trim()==AllowanceDetails[0]) && (table.getItem(i).getText_2(3).OleValue.toString().trim().indexOf(AllowanceDetails[1])!=-1) && (table.getItem(i).getText_2(6).OleValue.toString().trim()==AllowanceDetails[3])) { 
//      if((table.getItem(i).getText_2(5).OleValue.toString().trim()=="502")&&(table.getItem(i).getText_2(7).OleValue.toString().trim()=="1707100057")){
        flag = true;
        ApproveFlag = false;
          break;
      }
      else{ 
        table.Keys("[Tab]");
      }
      }
if(flag) { 
if(ImageRepository.ImageSet.Close_Filter.Exists()){ 
  ImageRepository.ImageSet.Close_Filter.Click();
  Delay(5000);
}


var Reject = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Management").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
Sys.HighlightObject(Reject);
Reject.Click();
Delay(3000);

var Reson = Sys.Process("Maconomy").SWTObject("Shell", "Reject Absence Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
Reson.Click();
if(AllowanceDetails[6]!="")
Reson.setText(AllowanceDetails[6]);
//else{ 
//  ValidationUtils.verify(false,true,"Reason is Needed to Reason Absence");
//  }
Delay(1000);
var Reject  = Sys.Process("Maconomy").SWTObject("Shell", "Reject Absence Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Reject Absence");
Reject.Click();
Delay(6000);


WorkspaceUtils.closeAllWorkspaces();


}

}
}


function gotTODOs_Approve(level){ 
  var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  Delay(4000);
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Delay(2000);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Delay(2000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(2000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(2000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(2000);
  Sys.Desktop.KeyDown(0x28);
  Sys.Desktop.KeyUp(0x28);
  Delay(2000);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  Delay(4000);
//  var refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
//  refresh.Click();
  
  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();

  
  
  Delay(15000);
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
if(level==0)
Client_Managt.DblClickItem("|Approve Absence (*)");

if(level==1)
Client_Managt.DblClickItem("|Approve Absence by Type (Substitute) (*)");


break;
}
}
}



}

function SOXexcel(CreateClient,start){ 
var Arrayss = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
     Arrayss[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrayss;
}

function RejectingAbsenceRequest(){ 
  gotoMenu();
  Absence();
  WorkspaceUtils.closeAllWorkspaces();
}