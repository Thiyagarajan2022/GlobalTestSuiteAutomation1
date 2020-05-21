﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "Reject Expenses";
  Indicator.Show();
  Indicator.PushText("waiting for window to open");

Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);
Log.Message(sheetName);
var Arrays = [];
var count = true;
var STIME = "";
var Description = "";
var Expense_Number = "";
var Approve_Level = [];
var y=0;
var w=0;
var ApproveInfo = [];
var logindetail = [];
var level =0;
var Language = "";
var comapany = "";
var approvers = [];
var OpCo2 = [];

var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "Reject Expenses", true);

function RejectExpense() {
  TextUtils.writeLog("Reject Expenses is Started");  
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
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
      Language = "";
      Language = EnvParams.Language;
        if((Language==null)||(Language=="")){
          ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
        }      
      Language = EnvParams.LanChange(Language);
      WorkspaceUtils.Language = Language;
      Log.Message(Language)
      
      excelName = EnvParams.path;
      workBook = Project.Path+excelName;
      STIME = "";
      STIME = WorkspaceUtils.StartTime();
      TextUtils.writeLog("Execution Start Time :"+STIME);
      Description;
      Expense_Number = "";
      Approve_Level = [];
      y=0;
      ApproveInfo = [];
      level =0; 
      logindetail = [];
      
      sheetName = "Reject Expenses";
      ExcelUtils.setExcelName(workBook, sheetName, true);
    goToJobMenuItem();
    Expense_Number = ExcelUtils.getRowDatas("Expense Number",EnvParams.Opco)
        if((Expense_Number=="")||(Expense_Number==null)){
              ExcelUtils.setExcelName(workBook, "Data Management", true);
              Expense_Number = ReadExcelSheet("Expense Number",EnvParams.Opco,"Data Management");
        } 
    
    gotoTimeExpenses();
    Allaprove();
    closeAllWorkspaces(); 
    
      CredentialLogin();
      for(var i=0;i<ApproveInfo.length;i++){
          WorkspaceUtils.closeMaconomy();
          aqUtils.Delay(10000, Indicator.Text);
          var temp = ApproveInfo[i].split("*"); 
          Restart.login(temp[2]);
          aqUtils.Delay(5000, Indicator.Text);
          todo(temp[3]);          
          rejtexpen(temp[0],temp[1],temp[2]);
      }
      closeAllWorkspaces();
}

function gotoTimeExpenses(){
    
    ReportUtils.logStep("INFO","Approve Expenses Second Level is Started:"+STIME);  
    aqUtils.Delay(2000,Indicator.Text); 
    Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Refresh(); 
    var expenses = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.expensestab;
    expenses.Click();
    expenses.HoverMouse();
    ReportUtils.logStep_Screenshot();
    aqUtils.Delay(1000,Indicator.Text);
    
//      ExcelUtils.setExcelName(workBook, sheetName, true);
//      Expense_Number = ExcelUtils.getRowDatas("Expense Number",EnvParams.Opco)
//        if((Expense_Number=="")||(Expense_Number==null)){
//              ExcelUtils.setExcelName(workBook, "Data Management", true);
//              Expense_Number = ReadExcelSheet("Expense Number",EnvParams.Opco,"Data Management");
//        }    
        

    
//    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh()
     var table = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
 
    var sheetno = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
    Sys.HighlightObject(sheetno);    
    sheetno.Click();
    sheetno.setText(Expense_Number);
    aqUtils.Delay(1000,Indicator.Text); 
  
    var flag=false;  
    for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(1).OleValue.toString().trim()==Expense_Number){ 
        flag=true;
        break;
      }
      else{ 
        table.Keys("[Down]");
      }
     }   
//     TextUtils.writeLog("Expense Sheet is available in Maconomy);
    ValidationUtils.verify(true,true,"Expense Sheet is available in Maconomy"); 
      
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
        
    }
    
    function Allaprove(){
        
        var Allaprovetab = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.PTabItemPanel.TabControl;
        var Add_Visible8 = true;
        while(Add_Visible8){
            if(Allaprovetab.isEnabled()){
              aqUtils.Delay(2000,Indicator.Text);
              Add_Visible8 = false;
              Allaprovetab.HoverMouse();
              ReportUtils.logStep_Screenshot();
              Allaprovetab.Click();
              aqUtils.Delay(2000,Indicator.Text);
              ImageRepository.ImageSet0.Maximize.Click();
        
              var All_Approver = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.TabFolderPanel.TabControl;
              aqUtils.Delay(1000,Indicator.Text);
              All_Approver.Click();
              aqUtils.Delay(3000,Indicator.Text);
              ReportUtils.logStep_Screenshot();
                    
                var Approval_table = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
                Sys.HighlightObject(Approval_table);               
                    for(var z=0;z<Approval_table.getItemCount();z++){ 
                        if(z<1){
                             approvers="";   
                             if(Approval_table.getItem(z).getText_2(8)!="Rejected"){      
                               approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
                               Approve_Level[y] = EnvParams.Opco+"*"+Expense_Number+"*"+approvers;
                               Log.Message(Approve_Level[y]);
                               ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
                               y++;
                             }  
                        }
                    }
          }
          var info_Bar = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.PTabItemPanel2.TabControl;
          info_Bar.Click();
          Delay(4000);
          ImageRepository.ImageSet0.Forward.Click();
          aqUtils.Delay(4000,Indicator.Text);
      
              var OpCo1 = EnvParams.Opco;
              ExcelUtils.setExcelName(workBook, "Server Details", true);
              var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
              var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
        }
    }


function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
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

function todo1(lvl){ 
   
    
      var toDo = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
    toDo.DBlClick();
    aqUtils.Delay(3000, Indicator.Text);
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
              if(lvl==2)
                Client_Managt.DblClickItem("|Approve Expense Sheet Line (*)");
              if(lvl==3)
                Client_Managt.DblClickItem("|Approve Expense Sheet Line (Substitute) (*)");
              break;
              }
          }
      }
}


function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
   
    var toDo = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.TabControl;
 
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
  
   var linestatus = false;
    if(!linestatus) 
    if((Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite3).isVisible())
    {
    var refresh = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
    linestatus = true;
    }
     if(!linestatus) 
    if((Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite2).isVisible())
    {   
    var refresh = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
    linestatus = true;
    }
  refresh.Click();
  aqUtils.Delay(15000, Indicator.Text);  
  
   var linestatus = false;
    if(!linestatus) 
    if((Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite3).isVisible())
    {
    var Client_Managt = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
    linestatus = true;
    }
     if(!linestatus) 
    if((Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite2).isVisible())
    {   
    var Client_Managt = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
    linestatus = true;
    }
   Sys.HighlightObject(Client_Managt)
    var listPass = true;
      if(lvl==2)
        for(var j=0;j<Client_Managt.getItemCount();j++){
          var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
          var temp1 = temp.split("(");
          if((temp.indexOf("Approve Expense Sheet Line (")!=-1)&&(temp1.length==2)){ 
            Client_Managt.ClickItem("|"+temp);   
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp);  
            TextUtils.writeLog("Entering into Approve Expenses Sheet Line from To-Dos List");
            listPass = false; 
          }
      }
      if(lvl==3)
      for(var j=0;j<Client_Managt.getItemCount();j++){ 
          var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
          var temp1 = temp.split("(");
        if((temp.indexOf("Approve Expense Sheet Line (Substitute) (")!=-1)&&(temp1.length==3)){ 
            Client_Managt.ClickItem("|"+temp);    
            ReportUtils.logStep_Screenshot(); 
            Client_Managt.DblClickItem("|"+temp); 
            TextUtils.writeLog("Entering into Approve Vendor by Type (Substitute) from To-Dos List");
            var listPass = true;   
         }
      }  
  if(listPass){
    if(lvl==2)
          for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
              if((temp.indexOf("Approve Expense Sheet Line by Type (")!=-1)&&(temp1.length==2)){ 
                  Client_Managt.ClickItem("|"+temp);   
                  ReportUtils.logStep_Screenshot(); 
                  Client_Managt.DblClickItem("|"+temp);  
                  TextUtils.writeLog("Entering into Approve Expenses Sheet Line by Type from To-Dos List");
                  listPass = false; 
                }
           }
    if(lvl==3)
        for(var j=0;j<Client_Managt.getItemCount();j++){ 
            var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
            var temp1 = temp.split("(");
          if((temp.indexOf("Approve Expense Sheet Line by Type (Substitute) (")!=-1)&&(temp1.length==3)){ 
              Client_Managt.ClickItem("|"+temp);    
              ReportUtils.logStep_Screenshot(); 
              Client_Managt.DblClickItem("|"+temp); 
              TextUtils.writeLog("Entering into Approve Expenses Sheet Line by Type (Substitute) from To-Dos List");
              var listPass = true;   
            }
        } 
  }
}


  
 function rejtexpen(company,Expense_Number,loginname){   
        var table = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder;
        waitForObj(table);
        
         if(Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Visible){
          }
          else{
          var showFilter = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
          waitForObj(table);
          Sys.HighlightObject(showFilter);
          showFilter.HoverMouse();
          showFilter.HoverMouse();
          showFilter.HoverMouse();
          showFilter.Click();
          }
          
          var table = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
          var firstCell = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
          waitForObj(firstCell);
          Sys.HighlightObject(firstCell);
          firstCell.HoverMouse();
          firstCell.Keys("[Tab]");
          var Expenseno = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
         
          aqUtils.Delay(3000, "Reading Data in table");;
          var closefilter = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
          waitForObj(closefilter);
          Sys.HighlightObject(closefilter);
          closefilter.HoverMouse();
          closefilter.HoverMouse(); 
          closefilter.HoverMouse();
          closefilter.HoverMouse();
          
            Expenseno.ClickM();
                table.Child(1).forceFocus();
                table.Child(1).setVisible(true);
                table.Child(1).setText("^a[BS]");
                table.Child(1).setText(Expense_Number);
                aqUtils.Delay(3000, Indicator.Text);
                var flag=false;
                  for(var v=0;v<table.getItemCount();v++){ 
                    if(table.getItem(v).getText_2(1).OleValue.toString().trim()==Expense_Number){ 
                      flag=true;
                      break;
                    }
                    else{ 
                      table.Keys("[Down]");
                    }
                  }
                TextUtils.writeLog("Expenses Sheet is listed for Reject");
                ValidationUtils.verify(flag,true,"Expenses Sheet is listed for Reject");
              Sys.HighlightObject(closefilter)    
             closefilter.Click();  
    
          var liness =  Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.TabControl;
         Sys.HighlightObject(liness);
         var lines = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
         Sys.HighlightObject(lines);
         var row =  lines.getItemCount();
         for(var l=1;l<row;l++){

                var lineapprove = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.PTabItemPanel.TabControl;
                 lineapprove.Click();
  
                 var lneaprove = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl;
                 Sys.HighlightObject(lneaprove);
                 lneaprove.Click();
                 aqUtils.Delay(1000,Indicator.Text);
 
                 var lneaprovetab = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
                Sys.HighlightObject(lneaprovetab);
                aqUtils.Delay(1000,Indicator.Text); 
                         
              var row = lneaprovetab.getItemCount();
              var col = lneaprovetab.getColumnCount();
   
              for(var i=0;i<row-2;i++){
                  var remark = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
                  remark.Click()
                  remark.setText("Rejected");
                  var save = Aliases.Maconomy.Expenses.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                 Sys.HighlightObject(save);
                  save.Click();  
                  TextUtils.writeLog("Reject reason is entered in the linelevel and saved");
                  aqUtils.Delay(1000, Indicator.Text);
                
                  if(l==1){
                      if(ImageRepository.ImageSet0.Reject.Exists()){
                        ImageRepository.ImageSet0.Reject.Click();
                        ReportUtils.logStep_Screenshot();    
                        TextUtils.writeLog("Created Expenses Linelevel is Rejected");                    
                        ValidationUtils.verify(true,true,"Created Expenses Linelevel:" +l+ " is Rejected by :"+loginname)
                      }
                      else{ 
                        ReportUtils.logStep("INFO","Reject Button Is Invisible");
                        Log.Warning(EnvParams.Opco+" - "+Expense_Number+" - Rejected :"+loginname);
                      } 
                  }
                  if(l>1){ 
                      if(ImageRepository.ImageSet0.reject01.Exists()){                        
                        ImageRepository.ImageSet0.reject01.Click();
                        ReportUtils.logStep_Screenshot();
                        ValidationUtils.verify(true,true,"Created Expenses Linelevel:" +l+ " is Rejected by :"+loginname)
                      }  
                      else{ 
                        ReportUtils.logStep("INFO","Reject Button Is Invisible");
                        Log.Warning(EnvParams.Opco+" - "+Expense_Number+" - Rejected :"+loginname);
                      } 
                  }                  
                
                  Sys.Desktop.KeyDown(0x09);
                  Sys.Desktop.KeyUp(0x09);
                  Sys.Desktop.KeyDown(0x09);
                  Sys.Desktop.KeyUp(0x09); 
                
                  aqUtils.Delay(2000, Indicator.Text);
                  ReportUtils.logStep_Screenshot();
                  ImageRepository.ImageSet0.Forward.Click(); 
                   aqUtils.Delay(2000, Indicator.Text);
                   lines.Click();
                   lines.HoverMouse();
                   ReportUtils.logStep_Screenshot();
                   aqUtils.Delay(2000, Indicator.Text);
                   Sys.Desktop.KeyDown(0x28);
                   Sys.Desktop.KeyUp(0x28);
               }    
         
         }         
//         AfterAllaprove();
  }
  
  
     function AfterAllaprove(){
        
        var Allaprovetab = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.allapproveactions;
        var Add_Visible8 = true;
        while(Add_Visible8){
            if(Allaprovetab.isEnabled()){
              aqUtils.Delay(2000,Indicator.Text);
              Add_Visible8 = false;
              Allaprovetab.HoverMouse();
              ReportUtils.logStep_Screenshot();
              Allaprovetab.Click();
              aqUtils.Delay(2000,Indicator.Text);
              ImageRepository.ImageSet0.Maximize.Click();
        
              var All_Approver = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.posting;
              aqUtils.Delay(1000,Indicator.Text);
              All_Approver.Click();
              aqUtils.Delay(3000,Indicator.Text);
              ReportUtils.logStep_Screenshot();
                    
                var Approval_table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite4.McClumpSashForm.Composite.Composite.McTableWidget.approvetable;
                Sys.HighlightObject(Approval_table);               
                    for(var z=0;z<Approval_table.getItemCount();z++){ 
                      if(z<1){
                             approvers="";   
                             if(Approval_table.getItem(z).getText_2(8)!="Rejected"){      
                              
                             }  
                      }
                    }
          }
          }
          }

function goToJobMenuItem(){
     var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
      ReportUtils.logStep_Screenshot("");
       menuBar.DblClick();
          if(ImageRepository.ImageSet01.TE.Exists())
          {
           ImageRepository.ImageSet01.TE.Click();// GL
          }
          else if(ImageRepository.ImageSet01.TE1.Exists())
          {
           ImageRepository.ImageSet01.TE1.Click();
          }
          else{
           ImageRepository.ImageSet01.TE2.Click();
          }   
    aqUtils.Delay(3000, Indicator.Text);
    Sys.Desktop.KeyDown(0x12);
    Sys.Desktop.KeyDown(0x20);
    Sys.Desktop.KeyUp(0x12);
    Sys.Desktop.KeyUp(0x20);
    Sys.Desktop.KeyDown(0x58);
    Sys.Desktop.KeyUp(0x58);  
    aqUtils.Delay(1000, Indicator.Text);
    var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
    var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
    Delay(3000);
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
        Client_Managt.ClickItem("|Time & Expenses");
        ReportUtils.logStep_Screenshot();
        Client_Managt.DblClickItem("|Time & Expenses");
      }
    }
    Delay(3000);
  }
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  //    try{
//    var refresh = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite5.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//    }
//    catch(e){
//    var refresh = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite5.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
//    }
//    refresh.Click();
//    aqUtils.Delay(15000, Indicator.Text);
//    try{
//    Client_Managt = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite5.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
//    }
//    catch(e){
//    Client_Managt = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite5.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
//    }
//    if(EnvParams.Country.toUpperCase()=="INDIA")
//       Runner.CallMethod("IND_ApprovePurchaseOrder.todo",lvl);
//    else{
//    if(lvl==3){
//    Client_Managt.ClickItem("|Approve Expense Sheet Line (Substitute) (*)");
//    ReportUtils.logStep_Screenshot(); 
//    Client_Managt.DblClickItem("|Approve Expense Sheet Line (Substitute) (*)");
//    }
//    if(lvl==2){
//    Client_Managt.ClickItem("|Approve Expense Sheet Line (*)");
//    ReportUtils.logStep_Screenshot(); 
//    Client_Managt.DblClickItem("|Approve Expense Sheet Line (*)");
//    }
//    }