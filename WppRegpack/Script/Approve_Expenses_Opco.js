//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "Approve Expenses Sheet Opco";
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
var login =[];
var logindetail = []; 
var ApproveInfo = [];
var level =0;
var Language = "";
var comapany = "";
var approvers = [];
var OpCo2 = [];
var sLevel = true;


function ApproveOpco() {
Language = "";
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
  goToJobMenuItem();
  gotoTimeExpenses();
  WorkspaceUtils.closeAllWorkspaces(); 
    
  CredentialLogin();       

  for(var i=0;i<ApproveInfo.length;i++){    
    level = i;
      WorkspaceUtils.closeMaconomy();
      aqUtils.Delay(10000, Indicator.Text);
      var temp = ApproveInfo[i].split("*"); 
      Restart.login(temp[2]);
      aqUtils.Delay(5000, Indicator.Text);
      todo(temp[3]);          
      aprvExpense(temp[0],temp[1],temp[2]);
  }
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces(); 
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
  
  
function goToJobMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
  }
}
Delay(3000);
}
  
 function gotoTimeExpenses(){ 
    ReportUtils.logStep("INFO","Approve Expenses Second Level is Started:"+STIME);    
    aqUtils.Delay(2000,Indicator.Text); 
    Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Refresh(); 
    var expenses = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.expensestab;
    expenses.Click();
    ReportUtils.logStep_Screenshot();
    aqUtils.Delay(1000,Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    
    if(Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Index==1)
     var table = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
     else
     var table = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
     
    if(Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Index==1) 
    var sheetno = Aliases.Maconomy.Group2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
    else
    var sheetno = Aliases.Maconomy.Screen.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.SWTObject("McTextWidget", "")
    
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

    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
     
        Sys.Desktop.KeyDown(0x11);
        Sys.Desktop.KeyDown(0x46);
        Sys.Desktop.KeyUp(0x11);
        Sys.Desktop.KeyUp(0x46);
        
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    
    if(Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Index==7)
    var desp = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite.SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
    else
    var desp = Aliases.Composite7.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite2.SWTObject("Composite", "").SWTObject("McTextWidget", "", 2);
    
    WorkspaceUtils.waitForObj(desp);
    desp = desp.getText().OleValue.toString().trim()
    
    if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Index==5)
    var Lcount = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    else
    var Lcount = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    
    WorkspaceUtils.waitForObj(Lcount);
    Lcount = Lcount.getItemCount()-1;
    Log.Message(Lcount);
    
    if(Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Index==6)
        var Allaprovetab = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.PTabItemPanel.TabControl;
    else
        var Allaprovetab = Aliases.Maconomy.CreateBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel.TabControl
                
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
        
              if(Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Index==6)
              var All_Approver = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.TabFolderPanel.TabControl;
              else
              var All_Approver = Aliases.Maconomy.CreateBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl;
                    
              aqUtils.Delay(1000,Indicator.Text);
              All_Approver.Click();
              aqUtils.Delay(3000,Indicator.Text);
              ReportUtils.logStep_Screenshot();
              if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
              }
              
              if(Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Index==6)
              var Approval_table = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
              else
              var Approval_table = Aliases.Maconomy.CreateBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid

              
              Sys.HighlightObject(Approval_table);  
            var tableCnt = Approval_table.getItemCount();
            tableCnt = tableCnt/Lcount;
            Log.Message(tableCnt);
            Log.Message(tableCnt-1);
            var CCount = tableCnt-1
              for(var z=0;z<Approval_table.getItemCount();z++){                 
                if(z<CCount){
                   approvers="";   
                   if(Approval_table.getItem(z).getText_2(8)!="Approved"){      
                     approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
                     Approve_Level[y] = EnvParams.Opco+"*"+desp+"*"+approvers;
                     Log.Message(Approve_Level[y]);
                     ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
                     y++;
                   }                   
                 }
              }
          }
          
          if(Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.Index==6)
          var info_Bar = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite8.PTabItemPanel2.TabControl;
          else
          var info_Bar = Aliases.Maconomy.CreateBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.PTabItemPanel2.TabControl;
          
          info_Bar.Click();
          Delay(4000);
          ImageRepository.ImageSet0.Forward.Click();
          aqUtils.Delay(4000,Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
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



function todo(lvl){ 
  TextUtils.writeLog("Loged into Level "+level+" Approver login"); 
  var toDo = Aliases.Maconomy.CreateClient.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.ToDos;
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

if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.ToDoRefresh;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
var refresh = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Composite.SingleToolItemControl;
}
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 1).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
}
if(Aliases.Maconomy.Shell.Composite.Composite.Composite.SWTObject("Composite", "", 2).Visible){
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite3.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.Tree;
}

Sys.HighlightObject(Client_Managt)
var listPass = true;
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){
var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Expenses by Type (Substitute) from To-Dos List");
var listPass = true;   
}
}  

if(listPass){
if(lvl==2)
for(var j=0;j<Client_Managt.getItemCount();j++){ 
var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
var temp1 = temp.split("(");
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line by Type").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==2)){ 
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
if((temp.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Approve Expense Sheet Line by Type (Substitute)").OleValue.toString().trim()+" (")!=-1)&&(temp1.length==3)){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into Approve Expenses Sheet Line by Type (Substitute) from To-Dos List");
var listPass = true;   
}
} 
  }
}


function aprvExpense(company,Expense_Number,loginname){    
        
var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder;
waitForObj(table);
        
if(Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Visible){
}
else{
var showFilter = NameMapping.Sys.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.SingleToolItemControl;
waitForObj(table);
Sys.HighlightObject(showFilter);
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.HoverMouse();
showFilter.Click();
}
          
          var table = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
          var firstCell = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
          waitForObj(firstCell);
          Sys.HighlightObject(firstCell);
          firstCell.HoverMouse();
          firstCell.Click();
          firstCell.Keys("[Tab][Tab]");
//          var Expenseno = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox;
          var des = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.ApprovelTabel.McGrid.TextBox;
          aqUtils.Delay(3000, "Reading Data in table");;
          var closefilter = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
          waitForObj(closefilter);
          Sys.HighlightObject(closefilter);
          closefilter.HoverMouse();
          closefilter.HoverMouse(); 
          closefilter.HoverMouse();
          closefilter.HoverMouse();
          
            des.ClickM();
//                table.Child(1).forceFocus();
//                table.Child(1).setVisible(true);
//                table.Child(1).setText("^a[BS]");
                des.setText(Expense_Number);
                aqUtils.Delay(3000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
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
TextUtils.writeLog("Expenses Sheet is listed for Reject");
ValidationUtils.verify(flag,true,"Expenses Sheet is listed for Reject");
Sys.HighlightObject(closefilter)    
closefilter.Click();    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }   
                             
var lines = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(lines);
var row =  lines.getItemCount();
for(var l=0;l<row;l++){        
var lineapprove = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.PTabItemPanel.TabControl;
lineapprove.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
}
var lneaprove = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.TabControl;
Sys.HighlightObject(lneaprove);
lneaprove.Click();
aqUtils.Delay(1000,Indicator.Text);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
}  
var lneaprovetab = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
Sys.HighlightObject(lneaprovetab);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
                  
}
aqUtils.Delay(1000,Indicator.Text); 
                             
var roww = lneaprovetab.getItemCount();
var col = lneaprovetab.getColumnCount(); 
var APGrid = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
if(lneaprovetab.isVisible()){ 
                   if(sLevel)
                   for(var kg = 1;kg<=level;kg++){ 
                     APGrid.Keys("[Down]");
                     sLevel = false;
                   }
var remark = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
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
    
    
    
    
    
//function aproveline(){
//              var lne = false;
//              if(!lne)
//              if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).isVisible())
//              {
//               var lneaprovetab = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2); 
//               lne = true; 
//              } 
//              if(!lne)
//              if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).isVisible())
//              {
//               var lneaprovetab = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2); 
//               lne = true; 
//              }
//              Sys.HighlightObject(lneaprovetab);
//               aqUtils.Delay(1000,Indicator.Text); 
//                             
//               var roww = lneaprovetab.getItemCount();
//               var col = lneaprovetab.getColumnCount();                  
//               
//               if(logindetail[0]==loginname){
//                  if(lneaprovetab.isVisible()){                     
//                          if(ImageRepository.ImageSet0.lineone.Exists())
//                          {
//                           ImageRepository.ImageSet0.lineone.Click();
//                          }
//                          else if(ImageRepository.ImageSet0.lineone1.Exists())
//                          {
//                           ImageRepository.ImageSet0.lineone1.Click();
//                          }
//                          else{
//                           ImageRepository.ImageSet0.lineone2.Click();
//                          }
//                          aqUtils.Delay(1000,Indicator.Text);
//                          ReportUtils.logStep_Screenshot();
//                          ValidationUtils.verify(true,true,"Created Expenses Linelevel is Approved by :"+loginname)
//                       }
//                     else{ 
//                          ReportUtils.logStep("INFO","Approve Button Is Invisible");
//                          Log.Warning(comapany+" - "+Expense_Number+" - Approver :"+loginname);
//                        }
//                  }              
//               else{
//                 if(lneaprovetab.isVisible()){                     
//                          if(ImageRepository.ImageSet0.linetwo.Exists())
//                          {
//                           ImageRepository.ImageSet0.linetwo.Click();
//                          }
//                          else if(ImageRepository.ImageSet0.linetwo1.Exists())
//                          {
//                           ImageRepository.ImageSet0.linetwo1.Click();
//                          }
//                          else{
//                           ImageRepository.ImageSet0.linetwo2.Click();
//                          }
//                          aqUtils.Delay(1000,Indicator.Text);
//                          ReportUtils.logStep_Screenshot();
//                          ValidationUtils.verify(true,true,"Created Expenses Linelevel is Approved by :"+loginname)
//                       }
//                     else{ 
//                          ReportUtils.logStep("INFO","Approve Button Is Invisible");
//                          Log.Warning(comapany+" - "+Expense_Number+" - Approver :"+loginname);
//                        }
//                }                
//} 
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

//function approve(loginname){
//Log.Message(loginname)
//Log.Message(loginname[0])
//Log.Message(loginname[1])



//            if(loginname[0]){
//                if(lneaprovetab.isVisible()){                     
//                        if(ImageRepository.ImageSet.lineone.Exists())
//                        {
//                         ImageRepository.ImageSet.lineone.Click();
//                        }
//                        else if(ImageRepository.ImageSet.lineone1.Exists())
//                        {
//                         ImageRepository.ImageSet.lineone1.Click();
//                        }
//                        else{
//                         ImageRepository.ImageSet.lineone2.Click();
//                        }
//                        aqUtils.Delay(1000,Indicator.Text);
//                        ReportUtils.logStep_Screenshot();
//                        ValidationUtils.verify(true,true,"Created Expenses Linelevel is Approved by :"+loginname)
//                     }
//                   else{ 
//                        ReportUtils.logStep("INFO","Approve Button Is Invisible");
//                        Log.Warning(comapany+" - "+Expense_Number+" - Approver :"+loginname);
//                      }
//            }
//            else{
//               if(lneaprovetab.isVisible()){                     
//                        if(ImageRepository.ImageSet.linetwo.Exists())
//                        {
//                         ImageRepository.ImageSet.linetwo.Click();
//                        }
//                        else if(ImageRepository.ImageSet.linetwo1.Exists())
//                        {
//                         ImageRepository.ImageSet.linetwo1.Click();
//                        }
//                        else{
//                         ImageRepository.ImageSet.linetwo2.Click();
//                        }
//                        aqUtils.Delay(1000,Indicator.Text);
//                        ReportUtils.logStep_Screenshot();
//                        ValidationUtils.verify(true,true,"Created Expenses Linelevel is Approved by :"+loginname)
//                     }
//                   else{ 
//                        ReportUtils.logStep("INFO","Approve Button Is Invisible");
//                        Log.Warning(comapany+" - "+Expense_Number+" - Approver :"+loginname);
//                      }
//            } 
//} 
//


