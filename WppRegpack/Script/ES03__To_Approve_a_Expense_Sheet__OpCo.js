﻿//USEUNIT ActionUtils
//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ObjectUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils


/** 
 * This script reject the created timesheet
 * @author  : Muthu Kumar M
 * @version : 3.0
 * Modified Date(MM/DD/YYYY) : 01/11/2022
 */
 
 
 
var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "Approve Expenses Sheet Opco";
Indicator.Show();
Indicator.PushText("waiting for window to open");


ExcelUtils.setExcelName(workBook, sheetName, true);
var Arrays = [];
var count = true;
var STIME = "";
var Description = "";
var Expense_Number = "";
var Approve_Level = [];
var y=0;
var w=0;
var login =[];
var logindetail = []; 
var ApproveInfo = [];
var level =0;
var Language = "";
var comapany = "";
var approvers = [];
var OpCo2 = [];
var sLevel = true;
var Maconomy_ParentAddress,Maconomy_Index = "";

//Main Function
function ApproveOpco() {

//Getting Language from EnvParamaters.xlsx
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

//Checking Login to execute Approve Expence by Opco Logins only
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);

var Macscreen = WorkspaceUtils.switch_Maconomy(Project_manager)
if(Macscreen=="Screen Not Found"){
Restart.login(Project_manager);
aqUtils.Delay(5000, Indicator.Text);
}else{ 
Maconomy_ParentAddress =   eval(Macscreen)
}

Sys.Refresh();
aqUtils.Delay(15000, Indicator.Text);
//ExcelUtils.setExcelName(workBook, "Agency Users", true);
//Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco);
Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(Project_manager);

//Re-Intialize Variable
STIME = WorkspaceUtils.StartTime();
excelName = EnvParams.path;
workBook = Project.Path+excelName;
STIME = "";
Description;
Expense_Number = "";
Approve_Level = [];
y=0;
ApproveInfo = [];
level =0; 
logindetail = [];    
sLevel = true;

getDetails();
goTo_TimeSheet();  
gotoExpenses();
Allaprove();
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
Workspace_Client.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);
    
ApproveInfo = WorkspaceUtils.CredentialLogin(Approve_Level,excelName);


  for(var i=0;i<ApproveInfo.length;i++){    
    level = i;
    
    
var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
Workspace_Client.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);

Maconomy_Index = WorkspaceUtils.Maconomy_Parent;
Maconomy_Index++;
WorkspaceUtils.Maconomy_Parent = Maconomy_Index;
Log.Message(Maconomy_Index);
    
// Restarting maconomy with Approver Logins
var temp = ApproveInfo[i].split("*");
var Macscreen = WorkspaceUtils.switch_Maconomy(temp[2])
if(Macscreen=="Screen Not Found"){
Restart.login(temp[2]);
}
aqUtils.Delay(5000, Indicator.Text);

Maconomy_ParentAddress = WorkspaceUtils.switch_Maconomy(temp[2]);
//Refreshing To-Do's List to find Submitted Job
//Refreshing To-Do's List to find Submitted Job
ActionUtils.ToDos_Selection(Maconomy_ParentAddress, level, temp[3], JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line").OleValue.toString().trim(),
JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line by Type").OleValue.toString().trim(),
JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line (Substitute)").OleValue.toString().trim(),
JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line by Type (Substitute)").OleValue.toString().trim())
                 
      aprvExpense(temp[0],temp[1],temp[2]);
  }

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
Workspace_Client.Click();
WorkspaceUtils.closeAllWorkspaces();
aqUtils.Delay(10000, Indicator.Text);

}


function getDetails(){
  sheetName = "Approve Expenses Sheet Opco";
ExcelUtils.setExcelName(workBook, "Data Management", true);
Expense_Number = ReadExcelSheet("Expense Number",EnvParams.Opco,"Data Management");
if((Expense_Number=="")||(Expense_Number==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Expense_Number = ExcelUtils.getRowDatas("Expense Number",EnvParams.Opco)
} 
if((Expense_Number=="")||(Expense_Number==null)){
 ValidationUtils.verify(true,false,"Expense Number is need to reject expense") 
}
} 
  
  

// Navigating to Time & Expenses from Time & Expenses Menu
function goTo_TimeSheet(){

var Workspace_Client = ObjectUtils.Workspace_Client_Object(Maconomy_ParentAddress);
ActionUtils.DoubleClick_with_Screenshot(Workspace_Client)


ActionUtils.Select_timesheet_from_workspace(); //Select Timesheet & Expense Image from workspace CLient
ActionUtils.Moving_intoWorkspace(Maconomy_ParentAddress,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());

}



function gotoExpenses(){
    
  ReportUtils.logStep("INFO", "Enter Expenses Details");
waitUntil_MaconomyScreen_loaded_Completely();
var ExpenseTab = getObjectAddress_JavaClasssName_Index_withTabText(Maconomy_ParentAddress,"TabControl",5,"Expenses");
  waitForObj(ExpenseTab)
  ReportUtils.logStep_Screenshot("");
  ExpenseTab.Click();

  waitUntil_MaconomyScreen_loaded_Completely();
  var AllExpenses = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"Button","All Expense Sheets");
  AllExpenses.Click();
  waitUntil_MaconomyScreen_loaded_Completely();
    

    
var table = getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,3);
var sheetno = ""
var childcount = 0;
var Add = [];



Sys.HighlightObject(table)
//sheetno = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McTextWidget",1)
sheetno = table.SWTObject("McTextWidget", "");
//sheetno = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
Sys.HighlightObject(sheetno)
    Log.Message(sheetno.FullName)
    Sys.HighlightObject(sheetno);    
    sheetno.Click();
    sheetno.setText(Expense_Number);
    aqUtils.Delay(1000,Indicator.Text); 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
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
    ValidationUtils.verify(flag,true,"Expense Sheet is available in Maconomy"); 
      
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
        
    }
    
    function Allaprove(){
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
//    var desp = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
 var desp = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McTextWidget",2);   
 
 
    Sys.HighlightObject(desp)
    Log.Message(desp.FullName)   
    desp.Click();
    WorkspaceUtils.waitForObj(desp);
    desp = desp.getText().OleValue.toString().trim()
    
 
    
    
    var Lcount = "";   
    Lcount = getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,1);
    Sys.HighlightObject(Lcount)

    Log.Message(Lcount.FullName) 
    WorkspaceUtils.waitForObj(Lcount);
    Lcount = Lcount.getItemCount()-1;
    Log.Message(Lcount);
     
     var Allaprovetab ;
  PropArray = new Array("JavaClassName", "Index","ChildCount","Visible");
  ValuesArray = new Array("PTabItemPanel", "3","1",true);
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  let objHeight = 1000;
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists)&&(obj[i_count].Parent.Left>0)){
    if(objHeight>obj[i_count].Parent.Height)
    Allaprovetab = obj[i_count];  
    objHeight = obj[i_count].Parent.Height;
  }
}
Allaprovetab = Allaprovetab.SWTObject("TabControl", "");  
Sys.HighlightObject(Allaprovetab);
//approve_bar.Click();

     Sys.HighlightObject(Allaprovetab)
        Log.Message(Allaprovetab.FullName)     

        var Add_Visible8 = true;
        while(Add_Visible8){
            if(Allaprovetab.isEnabled()){
              aqUtils.Delay(2000,Indicator.Text);
              Add_Visible8 = false;
              Allaprovetab.HoverMouse();
              ReportUtils.logStep_Screenshot();
              Allaprovetab.Click();
              
              waitUntil_MaconomyScreen_loaded_Completely();
              
              ImageRepository.ImageSet0.Maximize.Click();
        
      
var All_Approver;
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("TabControl", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible)
  if(obj[i_count].text=="All Approval Actions"){
  Sys.HighlightObject(obj[i_count]);
  All_Approver = obj[i_count];
  break;
 }
}
Sys.HighlightObject(All_Approver);
        
       Sys.HighlightObject(All_Approver) ;
              Log.Message(All_Approver.FullName)  
              aqUtils.Delay(1000,Indicator.Text);
              All_Approver.Click();
              aqUtils.Delay(3000,Indicator.Text);
              ReportUtils.logStep_Screenshot();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
//                var Approval_table = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
                
     var Approval_table = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
     
        
       Sys.HighlightObject(Approval_table) ;
              Log.Message(Approval_table.FullName)                 
              Sys.HighlightObject(Approval_table);  
            var tableCnt = Approval_table.getItemCount();
            tableCnt = tableCnt/Lcount;
            Log.Message(tableCnt);
            Log.Message(tableCnt-1);
            var CCount = tableCnt-1
              for(var z=0;z<Approval_table.getItemCount();z++){                 
                if(z<CCount){
                   approvers="";   
                   if(Approval_table.getItem(z).getText_2(8)!=JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approved").OleValue.toString().trim()){      
                     approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
                     Approve_Level[y] = EnvParams.Opco+"*"+desp+"*"+approvers;
                     Log.Message(Approve_Level[y]);
                     ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
                     y++;
                   }                   
                 }
              }
              
          }
     
var ApprovalTableBar;
ApprovalTableBar = getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"TabControl",1,3)
Sys.HighlightObject(ApprovalTableBar);  
Sys.HighlightObject(ApprovalTableBar)
Log.Message(ApprovalTableBar.FullName)
ApprovalTableBar.Click(); 

waitUntil_MaconomyScreen_loaded_Completely();

          ImageRepository.ImageSet0.Forward.Click();
          
waitUntil_MaconomyScreen_loaded_Completely();
      

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




function aprvExpense(company,Expense_Number,loginname){    
      
waitUntil_MaconomyScreen_loaded_Completely();

waitUntil_MaconomyScreen_loaded_Completely();

var Filter = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Show Filter List");
if(Filter!=null)
{
Sys.HighlightObject(Filter);
Filter.HoverMouse();
Filter.Click();
}

waitUntil_MaconomyScreen_loaded_Completely();
var table = getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,3);

Sys.HighlightObject(table)
var firstCell = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McValuePickerWidget",1);
waitForObj(firstCell);
Sys.HighlightObject(firstCell);
firstCell.HoverMouse();
firstCell.Click();
firstCell.Keys("[Tab][Tab]");
var des = table.SWTObject("McTextWidget", "", 2);
//getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McTextWidget",2,2);
aqUtils.Delay(3000, "Reading Data in table");;

waitUntil_MaconomyScreen_loaded_Completely();
var closefilter = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
if(closefilter!=null)
{
Sys.HighlightObject(closefilter);
closefilter.HoverMouse();
}
          
des.ClickM();

des.setText(Expense_Number);
aqUtils.Delay(3000, Indicator.Text);
    
waitUntil_MaconomyScreen_loaded_Completely();

var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==Expense_Number){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
TextUtils.writeLog("Expenses Sheet is listed for Approve");
ValidationUtils.verify(flag,true,"Expenses Sheet is listed for Approve");
var closefilter = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"SingleToolItemControl","Close Filter List");
if(closefilter!=null)
{
Sys.HighlightObject(closefilter);
closefilter.HoverMouse();
}
Sys.HighlightObject(closefilter)    
closefilter.Click();  


waitUntil_MaconomyScreen_loaded_Completely();


                             
var lines = getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,"McGrid",2,1);
Sys.HighlightObject(lines);
var row =  lines.getItemCount();
for(var l=0;l<row;l++){        



     var Allaprovetab ;
  PropArray = new Array("JavaClassName", "Index","ChildCount","Visible");
  ValuesArray = new Array("PTabItemPanel", "3","1",true);
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  let objHeight = 0;
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Exists)&&(obj[i_count].Parent.Left>0)){
    if(objHeight<obj[i_count].Parent.Height)
    Allaprovetab = obj[i_count];  
    objHeight = obj[i_count].Parent.Height;
  }
}
Allaprovetab = Allaprovetab.SWTObject("TabControl", "");  
Sys.HighlightObject(Allaprovetab);
Allaprovetab.Click();
waitUntil_MaconomyScreen_loaded_Completely();



var LineApproval = getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,"TabControl","Line Approval");
Sys.HighlightObject(LineApproval);
LineApproval.Click();
waitUntil_MaconomyScreen_loaded_Completely();
 

//var remark = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McTextWidget",1);
var APGrid = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
var lneaprovetab = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
Sys.HighlightObject(lneaprovetab);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
                  
}
aqUtils.Delay(1000,Indicator.Text); 
                             
var roww = lneaprovetab.getItemCount();
var col = lneaprovetab.getColumnCount(); 


if(lneaprovetab.isVisible()){ 
  APGrid = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",2);
                   if(sLevel)
                   for(var kg = 1;kg<=level;kg++){ 
                     APGrid.Keys("[Down]");
                     sLevel = false;
                   }
var remark = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McTextWidget",1);
APGrid = remark.Parent;
remark.Click();
APGrid.HoverMouse();
ImageRepository.ImageSet0.linetwo2.Click();
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
                  
}
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(true,true,"Linelevel:"+loginname)
ValidationUtils.verify(true,true,"Created Expenses Linelevel is Approved by :"+loginname)
 }
else{ 
    ReportUtils.logStep("INFO","Approve Button Is Invisible");
    Log.Warning(comapany+" - "+Expense_Number+" - Approver :"+loginname);
  }
      
              
   ImageRepository.ImageSet0.Forward.Click(); 
   aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
                  
}
   lines.Click();
   lines.HoverMouse();
   ReportUtils.logStep_Screenshot();
   aqUtils.Delay(2000, Indicator.Text);
   Sys.Desktop.KeyDown(0x28);
   Sys.Desktop.KeyUp(0x28);          
}   

    }
    
    
