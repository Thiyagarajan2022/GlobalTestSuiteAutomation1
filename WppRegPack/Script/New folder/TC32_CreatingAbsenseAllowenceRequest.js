﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "CreateAbsenceAllowanceRequest";
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


function Allowance(){ 
  var AllowanceReq = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  AllowanceReq.Click();
  Delay(4000);
  var NewAllowance = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 5);
  NewAllowance.Click();
  Delay(3000);
  AllowanceDetails = SOXexcel(sheetName,1);
  var absenceType = Sys.Process("Maconomy").SWTObject("Shell", "Create Allowance Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
  if(AllowanceDetails[0]!=""){
  absenceType.Click();
  WorkspaceUtils.SearchByValue(absenceType,"Absence Type",AllowanceDetails[0]);
  }else{ 
  ValidationUtils.verify(false,true,"Absence Type is Needed to Create Allowance");
  }
  var EntryDate = Sys.Process("Maconomy").SWTObject("Shell", "Create Allowance Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McDatePickerWidget", "", 2);
  if(AllowanceDetails[1]!=""){
  EntryDate.Click();
  WorkspaceUtils.CalenderDateSelection(EntryDate,AllowanceDetails[1]);
  }
  var NumOfDays = Sys.Process("Maconomy").SWTObject("Shell", "Create Allowance Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("McTextWidget", "", 2);
  if(AllowanceDetails[2]!=""){
  NumOfDays.Click();
  NumOfDays.setText(AllowanceDetails[2]);
  }
  var Reason = Sys.Process("Maconomy").SWTObject("Shell", "Create Allowance Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("McTextWidget", "", 2);
  if(AllowanceDetails[3]!=""){
  Reason.Click();
  Reason.setText(AllowanceDetails[3]);
  }
  
  var create = Sys.Process("Maconomy").SWTObject("Shell", "Create Allowance Request").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
  Sys.HighlightObject(create);
  create.Click();
  Delay(4000);
  var all = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All");
  all.Click();
  Delay(4000);
  var date = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  if(AllowanceDetails[1]!=""){
  date.Click();
  date.setText(AllowanceDetails[1]);
  }
  date.Keys("[Tab]");
  var absenceType = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);;
  if(AllowanceDetails[0]!=""){
  absenceType.Click();
  absenceType.setText(AllowanceDetails[0]);
  }
  absenceType.Keys("[Tab]");
  var NumOfDays = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);; 
  if(AllowanceDetails[2]!=""){
  NumOfDays.Click();
  NumOfDays.setText(AllowanceDetails[2]);
  }
  NumOfDays.Keys("[Tab]");
  var Reason = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);;;
  if(AllowanceDetails[3]!=""){
  Reason.Click();
  Reason.setText(AllowanceDetails[3]);
  }  
  Delay(4000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  for(var i=0;i<table.getItemCount();i++){ 
   if((table.getItem(i).getText_2(0).OleValue.toString().trim()==AllowanceDetails[1]) && (table.getItem(i).getText_2(1).OleValue.toString().trim()==AllowanceDetails[0]) && (table.getItem(i).getText_2(3).OleValue.toString().trim()==AllowanceDetails[3])) { 
     Log.Message("Allowance is Successfully Created");
     var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 6);
     Sys.HighlightObject(submit);
     submit.Click();
     var approver = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
     Sys.HighlightObject(approver);
     approver.Click();
     break;
   }
   else{ 
     table.Keys("Down");
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

function CreateAbsenceAllowence(){ 
  gotoMenu();
  Allowance();
  WorkspaceUtils.closeAllWorkspaces();
}