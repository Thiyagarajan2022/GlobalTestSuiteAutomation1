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
var Project_manager,employeeNo = "";
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
//    employeeVendor_GlobalVendorInfo();   
//    companyVendorInfo();
    closeAllWorkspaces();
}
catch(err){ 
  Log.Message(err);
}
}


function getDetails(){

ExcelUtils.setExcelName(workBook,"Data Management", true);
userName = ExcelUtils.getRowDatas("Employee & Employee Vendor and User Name",EnvParams.Opco)
//userName = ExcelUtils.getRowDatas("UserCreation_UserName",EnvParams.Opco)
if((userName==null)||(userName=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);
userName = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
}
Log.Message(userName)
if((userName==null)||(userName=="")) 
ValidationUtils.verify(false,true,"UserName is Needed to Block an User");

ExcelUtils.setExcelName(workBook,"Data Management", true);
employeeNo = ExcelUtils.getRowDatas("Employee & Employee Vendor and User NO",EnvParams.Opco)
if((employeeNo==null)||(employeeNo=="")){ 
ExcelUtils.setExcelName(workBook, sheetName, true);  
employeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)
}
Log.Message(employeeNo)
if((employeeNo==null)||(employeeNo=="")){ 
ValidationUtils.verify(false,true,"Employee Number is Needed to Block an User");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
terminationDate = ExcelUtils.getRowDatas("TerminationDate",EnvParams.Opco)
if (terminationDate == "AUTOFILL")
  terminationDate = getSpecificDate(0)
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

emp_Number = employeeNo;
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(4000,"Maconomy loading data");

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(4000,"Maconomy loading data");

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



if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(4000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
aqUtils.Delay(4000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

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
  
var  closefilter = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Close Filter List").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
closefilter = w;
}
Log.Message(closefilter.FullName);
Sys.HighlightObject(closefilter);
closefilter.Click();
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
ReportUtils.logStep_Screenshot();
ValidationUtils.verify(flag,true,"Employee Created is available in system");


aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var  Users = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
Users = w;
}
Log.Message(Users.FullName);
Sys.HighlightObject(Users);
Users.Click();


aqUtils.Delay(3000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}


var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
waitForObj(firstcell);
firstcell.Keys(userName);
aqUtils.Delay(3000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
ReportUtils.logStep_Screenshot("");

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
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
var  closefilter = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Close Filter List").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
closefilter = w;
}
Log.Message(closefilter.FullName);
Sys.HighlightObject(closefilter);
closefilter.Click();

    ReportUtils.logStep_Screenshot();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
//  
    
var  Delete = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "User Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
Delete = w;
}
Log.Message(Delete.FullName);
Sys.HighlightObject(Delete);
Delete = Delete.FullName

Delete = Delete.substring(0,Delete.lastIndexOf("."));
Log.Message(Delete);
Delete = eval(Delete);


Sys.HighlightObject(Delete);
var w = Delete.FindChild("toolTipText", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Next (Ctrl+D)").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
Delete = w;
}
Log.Message(Delete.FullName);
Sys.HighlightObject(Delete);

    Sys.HighlightObject(Delete);
    Delete.Click();
    
aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var Deleteuser = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete User").OleValue.toString().trim()+" ").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Next").OleValue.toString().trim())
Deleteuser.Click();

aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Maconomy loading data");
var Yes = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - User Information").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
Yes.Click();



aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Maconomy loading data");

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - User Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var Yes = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - User Information").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim());
Yes.Click();

}

aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Maconomy loading data");




var Delete = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete").OleValue.toString().trim()).SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
Delete.Click();

aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

 var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete User").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var Deleteuser = Sys.Process("Maconomy").SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Delete User").OleValue.toString().trim()+" ").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
Deleteuser.Click();
aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }


var  refresh = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "User Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
refresh = w;
}
Log.Message(refresh.FullName);
Sys.HighlightObject(refresh);
refresh = refresh.FullName

refresh = refresh.substring(0,refresh.lastIndexOf("."));
Log.Message(refresh);
refresh = eval(refresh);

var w = refresh.FindChild("toolTipText", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Refresh (F5)").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
refresh = w;
}
Log.Message(refresh.FullName);
Sys.HighlightObject(refresh);

    Sys.HighlightObject(refresh);
    refresh.Click();
}

aqUtils.Delay(5000,"Maconomy loading data");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(5000,"Maconomy loading data");

var  showFilter = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Show Filter List").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
showFilter = w;
}
Log.Message(showFilter.FullName);
Sys.HighlightObject(showFilter);
showFilter.Click();

    ReportUtils.logStep_Screenshot();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(3000,"Maconomy loading data");
    
var  Users = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Users").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
Users = w;
}
Log.Message(Users.FullName);
Sys.HighlightObject(Users);
Users.Click();


aqUtils.Delay(3000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}



var  table = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("JavaClassName", "McFilterPaneWidget", 2000);
  if (w.Exists)
{ 
table = w;
}
table = table.SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
Log.Message(table.FullName);
Sys.HighlightObject(table);
waitForObj(table);
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}   
    var flag=true;
    for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userName)){ 
       flag=false;
       Log.Message("User is available");
       break;
      }else{ 
        table.Keys("[Down]");
      }
    }

  if(flag==false)
    ValidationUtils.verify(flag,true,"User is AVAILABLE in system");
  else
    ValidationUtils.verify(flag,true,"User is NOT AVAILABLE in system");
    
    

    
    if(flag){ 
      
    aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var  EmployeeVendor = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee Vendor").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
EmployeeVendor = w;
}
Log.Message(EmployeeVendor.FullName);
Sys.HighlightObject(EmployeeVendor);
EmployeeVendor.Click();
    aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}


var Screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
Screen.Refresh();
aqUtils.Delay(2000,"Loading Employee");  
//var  table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var  table = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("JavaClassName", "McFilterPaneWidget", 2000);
  if (w.Exists)
{ 
table = w;
}
table = table.SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
Log.Message(table.FullName);
Sys.HighlightObject(table);


//var  firstCell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");

var  firstCell = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("JavaClassName", "McFilterPaneWidget", 2000);
  if (w.Exists)
{ 
firstCell = w;
}
firstCell = firstCell.SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//var firstCell = firstCell.SWTObject("McValuePickerWidget", "");

firstCell.Click();
firstCell.Keys("[Tab][Tab][Tab]");
aqUtils.Delay(4000,"Loading Employee"); 
//var Emp_No = table.SWTObject("McValuePickerWidget", "");
//var  Emp_No = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");

var Screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
Screen.Refresh();
aqUtils.Delay(2000,"Loading Employee"); 
var  Emp_No = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("JavaClassName", "McFilterPaneWidget", 2000);
  if (w.Exists)
{ 
Emp_No = w;
}
Emp_No = Emp_No.SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
Emp_No.setText(emp_Number);

if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
  if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(emp_Number)){ 

    flag=true;    
    break;
  }
  else{ 
    table.Keys("[Down]");
  }
}

if(flag)
ValidationUtils.verify(true,true,"Employee Number is AVAILABLE in Maconomy");
else
ValidationUtils.verify(flag,true,"Employee Number is NOT AVAILABLE in Maconomy");

if(flag){ 
var  closefilter = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Close Filter List").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
closefilter = w;
}
Log.Message(closefilter.FullName);
Sys.HighlightObject(closefilter);
closefilter.Click();

    ReportUtils.logStep_Screenshot();
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}



ImageRepository.ImageSet.Maximize1.Click();

var screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "");
screen.Click();
screen.MouseWheel(-10);

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
var Blocked = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
//Blocked.Click();
  if(Blocked.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(true,true,"Employee Vendor is already BLOCKED");
  else{ 
  Blocked.Click();
  aqUtils.Delay(5000,"Maconomy loading data");
  DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(true,true,"Company Employee Vendor is BLOCKED");
}

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
var Status = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
 Status.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
  Status.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Inactive").OleValue.toString().trim(),"Status")

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
Sys.HighlightObject(Save);
Save.Click();

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
    
     var p = Sys.Process("Maconomy");
  Sys.HighlightObject(p);
  Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Company Vendor Information").OleValue.toString().trim())
 var w = p.FindChild("WndCaption", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employees - Company Vendor Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
  
var label = w.SWTObject("Label", "*");
Log.Message(label.getText());
var lab = label.getText().OleValue.toString().trim();
ReportUtils.logStep("INFO",lab)
TextUtils.writeLog(lab);
var Ok = w.SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
Ok.HoverMouse();
ReportUtils.logStep_Screenshot("");
Ok.Click();
}
    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
var Close_Company_Information = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 1).SWTObject("TabControl", "");
Sys.HighlightObject(Close_Company_Information);;
Close_Company_Information.Click();

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

var CloseInformation = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;;
Sys.HighlightObject(CloseInformation);
CloseInformation.Click();
   
    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
//var Screen = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", ""); 
var Screen = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10;
Screen.Click();
Screen.MouseWheel(-10);;


    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
var Blocked = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
//Blocked.Click();
  if(Blocked.getText()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(true,true,"Global Employee Vendor is already BLOCKED");
  else{ 
  Blocked.Click();
  aqUtils.Delay(5000,"Maconomy loading data");
  DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Yes").OleValue.toString().trim())
  ValidationUtils.verify(true,true,"Global Employee Vendor is BLOCKED");
}

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    
var Status = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);
 Status.Keys(" ");
 aqUtils.Delay(5000, Indicator.Text);
  Status.Click();
  WorkspaceUtils.DropDownList(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Inactive").OleValue.toString().trim(),"Status")

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
var Save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);;
Sys.HighlightObject(Save);
Save.Click();

    aqUtils.Delay(3000,"Maconomy loading data");
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    

var  EmployeeVendor = ""; 
var p = Sys.Process("Maconomy");
Sys.HighlightObject(p);
var w = p.FindChild("text", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee Information").OleValue.toString().trim(), 2000);
  if (w.Exists)
{ 
EmployeeVendor = w;
}
Log.Message(EmployeeVendor.FullName);
Sys.HighlightObject(EmployeeVendor);
EmployeeVendor.Click();

    aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}
    aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}

//  var  Screen = ""; 
//var p = Sys.Process("Maconomy");
//Sys.HighlightObject(p);
//var w = p.FindChild("JavaClassName", "McPaneGui$10", 2000);
//  if (w.Exists)
//{ 
//Screen = w;
//}


//var Termination_Date = Screen.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McDatePickerWidget", "", 2);
var Termination_Date = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 4).SWTObject("McDatePickerWidget", "", 2);
Termination_Date.setText(terminationDate);
aqUtils.Delay(2000,"Loading Employee");  

//var Blocked = Screen.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
var Blocked = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
    if(!Blocked.getSelection()){
      Blocked.Click();
      ValidationUtils.verify(true,true,"Employee is BLOCKED");

    } 
 
        aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}   

//  var  Screen = ""; 
//var p = Sys.Process("Maconomy");
//Sys.HighlightObject(p);
//var w = p.FindChild("JavaClassName", "PTabFolder", 2000);
//  if (w.Exists)
//{ 
//Screen = w;
//}

//var Save = Screen.SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
var Save = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite10.Composite.PTabFolder.TabFolderPanel.SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
Sys.HighlightObject(Save)
Save.Click();

        aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}   
        aqUtils.Delay(2000,"Loading Employee");  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){}   

    
    }
    
    }
    
    
  
    }   



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

 



 
