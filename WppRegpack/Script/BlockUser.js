//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "BlockUser";
var STIME="";
var Project_manager = "";
ExcelUtils.setExcelName(workBook, sheetName, true);

var emp_Number,userName,terminationDate;


/**
  *  This Main function invokes maconomy and calls subfunctionality methods
  */
function blockUser()
{
TextUtils.writeLog("Block New User"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(2000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;
ExcelUtils.setExcelName(workBook, "Agency Users", true);
Project_manager = ExcelUtils.getRowDatas("Agency - Finance",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Block User test started::"+STIME);
TextUtils.writeLog("Block User test started::"+STIME);

try{
    getDetails();
    //goToMenu_user();
    //searchforUser();
    //deleteUser();
    //closeAllWorkspaces();
    goToMenu_employee();
    searchForEmployee();
    employeeVendor_GlobalVendorInfo();   
    companyVendorInfo();
    closeAllWorkspaces();
}
catch(err){ 
  Log.Message(err);
}
}


function getDetails(){

ExcelUtils.setExcelName(workBook,"Data Management", true);
userName = ExcelUtils.getRowDatas("UserCreation_UserName",EnvParams.Opco)
if((userName==null)||(userName=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
userName = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
}
Log.Message(userName)
if((userName==null)||(userName=="")) 
ValidationUtils.verify(false,true,"UserName is Needed to Block an User");


ExcelUtils.setExcelName(workBook, sheetName, true);
terminationDate = ExcelUtils.getRowDatas("TerminationDate",EnvParams.Opco)
Log.Message(terminationDate)
if((terminationDate==null)||(terminationDate=="")) 
ValidationUtils.verify(false,true,"Termination Date is Needed to Block an User");


}
    

function goToMenu_user(){ 

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
  menuBar.DblClick();
Log.Message(Language) 
  if(Language == "Chinese (Simplified)"){ 
if(ImageRepository.ImageSet4.Emp_Chinese_1.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_1.Click();
}
else if(ImageRepository.ImageSet4.Emp_Chinese_2.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_2.Click();
}
else if(ImageRepository.ImageSet4.Emp_Chinese_3.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_3.Click();  
}
  }else if(Language == "Spanish"){ 
    
if(ImageRepository.ImageSet4.Emp_Spanish_1.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_1.Click();
}
else if(ImageRepository.ImageSet4.Emp_Spanish_2.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_2.Click();
}
else if(ImageRepository.ImageSet4.Emp_Spanish_3.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_3.Click();  
}

  }else{
if(ImageRepository.ImageSet.Human_Resource_1.Exists()){
ImageRepository.ImageSet.Human_Resource_1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();  
}
}

  aqUtils.Delay(1000, "Finding Users");
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, "Finding Users");
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users").OleValue.toString().trim());

}

}
TextUtils.writeLog("Entering into Users from Human Resources Menu");  
  
}
      

function searchforUser()
{
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
var active_users =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language,"Active Users").OleValue.toString().trim());
waitForObj(active_users);
ReportUtils.logStep_Screenshot("");
active_users.Click();

var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
waitForObj(firstcell);
firstcell.Keys(userName);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
ReportUtils.logStep_Screenshot("");

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
waitForObj(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}   
    var flag=false;
    for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userName)){ 
       flag=true;
       Log.Message("User is available");
       break;
      }else{ 
        table.Keys("[Down]");
      }
    }

  if(flag==false)
    ValidationUtils.verify(flag,true,"User is not available in system");
  else
    ValidationUtils.verify(flag,true,"User is available in system");
 ReportUtils.logStep_Screenshot("");
  
if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    closefilter.HoverMouse();
    ReportUtils.logStep_Screenshot();
    closefilter.Click();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    }   
}
    
function deleteUser()
{
aqUtils.Delay(1000,"Waiting till page loads");

var emp_NumberObj = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);

waitForObj(emp_NumberObj);
emp_Number = emp_NumberObj.getText().OleValue.toString().trim();

var deleteuserButton = Aliases.Maconomy.Composite.SingleToolItemControl2;
waitForObj(deleteuserButton);
deleteuserButton.Click();

var deleteUserwindow = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete User").OleValue.toString()+" ");
waitForObj(deleteUserwindow);

var nextButton = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete User").OleValue.toString()+" ").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Next").OleValue.toString().trim())
waitForObj(nextButton);
nextButton.Click();

//if((Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption== JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim())){
//var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim()).SWTObject("Label", "*").getText();
//Log.Message(label1);
//var cancel = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
//cancel.HoverMouse();
//ReportUtils.logStep_Screenshot();
//cancel.Click();
//aqUtils.Delay(1000,"Waiting for Next Popup Window");
//}

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users - User Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
  {  
    var label = w.SWTObject("Label", "*");
    Log.Message(label.getText());
    var lab = label.getText().OleValue.toString().trim();
    ReportUtils.logStep("INFO",lab)
    TextUtils.writeLog(lab);
    var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
    Ok.HoverMouse();
    ReportUtils.logStep_Screenshot("");
    Ok.Click();
}
aqUtils.Delay(10000,"Waiting for Next Popup Window");

if((Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption== JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete").OleValue.toString().trim())){
var label1 = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete").OleValue.toString().trim()).SWTObject("Label", "*").getText();
Log.Message(label1);
var Ok = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Ok.HoverMouse();
ReportUtils.logStep_Screenshot();
Ok.Click();
aqUtils.Delay(10000,"Waiting to load Main Screen");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
}

}
    


      
function goToMenu_employee(){ 
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
  menuBar.DblClick();
Log.Message(Language) 
  if(Language == "Chinese (Simplified)"){ 
if(ImageRepository.ImageSet4.Emp_Chinese_1.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_1.Click();
}
else if(ImageRepository.ImageSet4.Emp_Chinese_2.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_2.Click();
}
else if(ImageRepository.ImageSet4.Emp_Chinese_3.Exists()){
ImageRepository.ImageSet4.Emp_Chinese_3.Click();  
}
  }else if(Language == "Spanish"){ 
    
if(ImageRepository.ImageSet4.Emp_Spanish_1.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_1.Click();
}
else if(ImageRepository.ImageSet4.Emp_Spanish_2.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_2.Click();
}
else if(ImageRepository.ImageSet4.Emp_Spanish_3.Exists()){
ImageRepository.ImageSet4.Emp_Spanish_3.Click();  
}

  }else{
if(ImageRepository.ImageSet.Human_Resource_1.Exists()){
ImageRepository.ImageSet.Human_Resource_1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();  
}
}

  aqUtils.Delay(1000, "Finding Employees");
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, "Finding Employees");
var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
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
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees").OleValue.toString().trim());

}

}
TextUtils.writeLog("Entering into Employee from Human Resources Menu");

}




function searchForEmployee()
{ 

emp_Number = "128410526";
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
Log.Message("Employee Number:"+emp_Number);
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
waitForObj(table);

var firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
waitForObj(firstCell);
firstCell.Click();
firstCell.Keys("[Tab]");

var mainCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").ChildCount;
var main = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
Sys.HighlightObject(main);
var MainBrnch = "";
for(var bi=0;bi<mainCount;bi++){ 
  if((main.Child(bi).isVisible())&&(main.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(main.Child(bi).Child(0).isVisible())){ 
    MainBrnch = main.Child(bi);
    break;
  }
}
Log.Message(MainBrnch.FullName);

var empNumberObj = MainBrnch.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2)
empNumberObj.setText(emp_Number);

var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(1000,"Loading Employee");  
ReportUtils.logStep_Screenshot();
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(1).OleValue.toString().trim()==(emp_Number)){ 

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

if(flag){ 
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
}
      
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Employee Created is available in system");
  

//var mainCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).ChildCount;
//var main = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2);
//Sys.HighlightObject(main);
//var MainBrnch = "";
//for(var bi=0;bi<mainCount;bi++){ 
//  if((main.Child(bi).isVisible())&&(main.Child(bi).Child(0).Name.indexOf("PTabFolder")!=-1)&&(main.Child(bi).Child(0).isVisible())){ 
//    MainBrnch = main.Child(bi);
//    break;
//  }
//}
//Log.Message(MainBrnch.FullName);

//var terminationdateObj =MainBrnch.SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McDatePickerWidget", "", 2);

var terminationdateObj = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McDatePickerWidget", "", 2)
waitForObj(terminationdateObj);
WorkspaceUtils.CalenderDateSelection(terminationdateObj,terminationDate);
  
var blockedCheckbox =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
blockedCheckbox.Click();

ReportUtils.logStep_Screenshot();
aqUtils.Delay(1000,"Waiting for Save button to enable")    
savebutton = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3)
savebutton.Click();

var employeeVendorTab = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
waitForObj(employeeVendorTab);
employeeVendorTab.Click();
           
}
function employeeVendor_GlobalVendorInfo()
{
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var globalVendorInfoSection = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "")
waitForObj(globalVendorInfoSection);
globalVendorInfoSection.Click();
globalVendorInfoSection.MouseWheel(-500);

var blockedDropDown = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2)
waitForObj(blockedDropDown);
blockedDropDown.Click();
Sys.Process("Maconomy").Refresh(); 
WorkspaceUtils.DropDownList("Yes","Blocked"); 

var statusDropDown = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2)
waitForObj(statusDropDown);
statusDropDown.Click();
Sys.Process("Maconomy").Refresh(); 
WorkspaceUtils.DropDownList("Inactive","Status"); 

ReportUtils.logStep_Screenshot();

var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3)
waitForObj(save);
save.Click();
}


function companyVendorInfo()
{
  
var closeFilter_ListCompanyVendor = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "")
closeFilter_ListCompanyVendor.Click();


var companyVendorinfo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");
companyVendorinfo.Click();
companyVendorinfo.MouseWheel(-500);

ReportUtils.logStep_Screenshot();
var companyvendorInfoBlocked =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
waitForObj(companyvendorInfoBlocked);
if(companyvendorInfoBlocked.getText() == "Yes")
    ValidationUtils.verify(true,true,"company vendor is blocked");
else
    ValidationUtils.verify(false,true,"company vendor is not blocked");
    
    
var companyVendorInfoStatus = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
waitForObj(companyVendorInfoStatus);
if(companyVendorInfoStatus.getText() == "Inactive")
    ValidationUtils.verify(true,true,"company vendor status is Inative");
else
    ValidationUtils.verify(false,true,"company vendor status is Active");

}

 



 
