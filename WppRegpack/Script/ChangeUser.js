//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT Restart

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var userInfo = [];
Approve_Level = [];
y=0;
ApproveInfo = [];
level =0;
var UserPasswd = [];
var sheetName = "ChangeUser";
//var sheetName1 = "Access Level";
//var loginpassword = "Credentials";
//var sscCredential = "SSC Credential";
var Credential = "userRoles";
var third_lvl_approver = false;
var login_satuts;
var STIME = "";
var LoginArr = [];
var HRData = [];
var LoginEmp = [];
//var lastArray = [];
//var Arrays = [];
var LoginArrays = [];
ExcelUtils.setExcelName(workBook, sheetName, true);
var nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,comapany,AccessLvel;


function changeuser()
{
  Indicator.PushText("waiting for window to open");
//aqTestCase.Begin("Job Creation", "zfj://CH1-67");
TextUtils.writeLog("Change User Started"); 
excelName = EnvParams.getEnvironment();
workBook = Project.Path+excelName;
sheetName = "ChangeUser";
Language = "";
Log.Message(workBook);
Log.Message(excelName);
//ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
//Log.Message(workBook);
ExcelUtils.setExcelName(workBook, sheetName, true);

Arrays = [];
count = true;
checkmark = false;
STIME = "";
nameValue,employeeNo,userType,validFrom,validTo,companyNo,validateCompany,comapany="";

Language = EnvParams.Language;
if((Language==null)||(Language=="")){
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}

Language = EnvParams.LanChange(Language);
WorkspaceUtils.Language = Language;
    if(Language=="English"){
    STIME = WorkspaceUtils.StartTime();
    ReportUtils.logStep("INFO", "Change User started::"+STIME);

   goToMenu();
    getDetails();
    goToUsers();
    user_Approval();

  closeAllWorkspaces();
 
  CredentialLogin();

for(var i=level;i<ApproveInfo.length;i++){
WorkspaceUtils.closeMaconomy();
aqUtils.Delay(10000, Indicator.Text);

var temp = ApproveInfo[i].split("*");
Restart.login(temp[2]);
aqUtils.Delay(5000, Indicator.Text);


todo(temp[3]);
aprvUser(temp[0],temp[1],temp[2],temp[3],i);


  }
  WorkspaceUtils.closeAllWorkspaces();
}


function getDetails(){
  var sheetName = "ChangeUser";
ExcelUtils.setExcelName(workBook, sheetName, true);

nameValue = ExcelUtils.getRowDatas("Name",EnvParams.Opco)


if((nameValue==null)||(nameValue=="")){ 
ValidationUtils.verify(false,true,"Name is required to create USER");
}


employeeNo = ExcelUtils.getRowDatas("Employee No",EnvParams.Opco)

if((employeeNo==null)||(employeeNo=="")){ 
ValidationUtils.verify(false,true,"employeeNo is Needed to Create a USER");
}

//Log.Message(employeeNo);

userType = ExcelUtils.getRowDatas("User Type",EnvParams.Opco)
if((userType==null)||(userType=="")){ 
ValidationUtils.verify(false,true,"User Type is required to create USER");
}

//Log.Message("userType"+userType);

validFrom = ExcelUtils.getRowDatas("Valid From",EnvParams.Opco)
if((validFrom==null)||(validFrom=="")){ 
ValidationUtils.verify(false,true,"Valid From is required to create USER");
}
//Log.Message("validFrom"+validFrom);

validTo = ExcelUtils.getRowDatas("Valid To",EnvParams.Opco)
if((validTo==null)||(validTo=="")){ 
ValidationUtils.verify(false,true,"Valid To is required to create USER");
}



AccessLevel = ExcelUtils.getRowDatas("Access Level",EnvParams.Opco)

if((AccessLevel==null)||(AccessLevel=="")){ 
//ValidationUtils.verify(false,true,"Acc is Needed to Create a USER");
}

companyNo = ExcelUtils.getRowDatas("company",EnvParams.Opco)
if((companyNo==null)||(companyNo=="")){ 
ValidationUtils.verify(false,true,"CompanyNo is required to create USER");
}

//Log.Message("companyNo"+companyNo);

//validateCompany = getExcelData_Company("Validate_Company",EnvParams.Opco)
//if((validateCompany==null)||(validateCompany=="")){ 
//ValidationUtils.verify(false,true,"validateCompany is required to create USER");
//}
////Log.Message("validateCompany"+validateCompany);




ExcelUtils.setExcelName(workBook, "Server Details", true);
Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)

}

function goToMenu(){ 
 ReportUtils.logStep_Screenshot("Navigating to Menu"); 
 TextUtils.writeLog("Navigating to Menu"); 
    var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
    menuBar.DblClick();
    

     if(ImageRepository.ImageSet.HR.Exists()){
    ImageRepository.ImageSet.HR.Click();
    }
    else if(ImageRepository.ImageSet.HR1.Exists()){
    ImageRepository.ImageSet.HR1.Click();
    }
    else if(ImageRepository.ImageSet.HR2.Exists()){
    ImageRepository.ImageSet.HR2.Click();  
    }

  aqUtils.Delay(3000, Indicator.Text);
//  Delay(3000);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
//  Delay(2000);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);
var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
aqUtils.Delay(1000);
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}

      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
      var Client_Managt;
      for(var i=1;i<=childCC;i++){ 
      Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
      if(Client_Managt.isVisible()){ 
      Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
           Client_Managt.HoverMouse();
ReportUtils.logStep_Screenshot("");
      Client_Managt.DblClickItem("|Users");
      
      }
      }
      aqUtils.Delay(1000);
      Log.Message("User Icon is Clicked");
      }
      
      
      
function getExcelData_Company(rowidentifier,column) { 
excelData =[];  

var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
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

      
      
function goToUsers(){ 
aqUtils.Delay(1000);
ReportUtils.logStep("INFO", "Enter User Details");
ReportUtils.logStep_Screenshot("");
     aqUtils.Delay(1000);
    }
    
    
    function test()
    {
      
        var periodfrom = 
 Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2) 
      Sys.HighlightObject(periodfrom);
    }
    
    function user_Approval(){ 
 aqUtils.Delay(1000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
 firstcell.Keys(nameValue);
 ReportUtils.logStep_Screenshot(""); 
 TextUtils.writeLog("Searching For User"); 
     ReportUtils.logStep_Screenshot("");
  firstcell.Keys("[Tab][Tab]");
  aqUtils.Delay(1000);

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  aqUtils.Delay(1000);  
    var flag=false;
    for(var v=0;v<table.getItemCount();v++){ 
  //    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userInfo[0]+" "+STIME))
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()=="Muthu 18:42:32"){ 
        flag=true;
        Log.Message("User is available");
        break;
      }else{ 
        table.Keys("[Down]");
      }
    }

//    var flag = table.getItemCount()>0;
    ValidationUtils.verify(flag,true,"User is available in system");
       ReportUtils.logStep_Screenshot("");
  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
   
    closefilter.HoverMouse();
    ReportUtils.logStep_Screenshot();
    closefilter.Click();
    aqUtils.Delay(1000);

    }
    aqUtils.Delay(1000);
    if((AccessLevel!="")&&(AccessLevel!=null)){ 
      
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    var Access_Level = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
    Sys.HighlightObject(Access_Level);
       Access_Level.HoverMouse();
    ReportUtils.logStep_Screenshot();
    Access_Level.Click();
    aqUtils.Delay(1000);
    
  var add = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
     add.HoverMouse();
    ReportUtils.logStep_Screenshot(); 
  add.Click();
  aqUtils.Delay(1000);
  
   ReportUtils.logStep_Screenshot(""); 
 TextUtils.writeLog(""); 
  
  var cell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//  SearchByValues_Col_1(ObjectAddrs,popupName,value,fieldName)

   cell.Click();
   aqUtils.Delay(1000);
   WorkspaceUtils.AccessLevel_Add(cell,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Access Level").OleValue.toString().trim(),AccessLevel,"AccessLevel");

  var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);;
 
      save.HoverMouse();
    ReportUtils.logStep_Screenshot(); 
  save.Click();
}
aqUtils.Delay(1000);

   ReportUtils.logStep_Screenshot(""); 
 TextUtils.writeLog(""); 

  if((validFrom!="")&&(validFrom!=null)){ 
    
    var periodfrom = 
 Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 2)   
  if(validFrom!=""){
    WorkspaceUtils.CalenderDateSelection(periodfrom,validFrom)
     ReportUtils.logStep_Screenshot("");
     }
  }
  
    if((validTo!="")&&(validTo!=null)){ 
    
  var periodto = 
  Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McDatePickerWidget", "", 4)
  if(validTo!=""){
  WorkspaceUtils.CalenderDateSelection(periodto,validTo)

     ReportUtils.logStep_Screenshot("");
  }  
  
  }
  
     ReportUtils.logStep_Screenshot(""); 
 TextUtils.writeLog(""); 
 
     aqUtils.Delay(1000);
var savebtn =""

if((Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").Index==2)  &&
(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible()))
savebtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3)
  
else if((Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==2)  &&
(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible()))
savebtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3)  

aqUtils.Delay(1000);
     Sys.HighlightObject(savebtn);
     savebtn.HoverMouse();
     savebtn.Click();
    ReportUtils.logStep_Screenshot(); 

 TextUtils.writeLog("Saving the changes"); 
  
 
    
aqUtils.Delay(4000);

if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Users - User Information")    
{
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - User Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - User Information").SWTObject("Label", "*").WndCaption;
      Log.Message(label );
       button.HoverMouse();
   //  ReportUtils.logStep_Screenshot("");
      button.Click();
      aqUtils.Delay(1000);
     } 
     
     

if(Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption=="Users - User Information")    
{
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - User Information").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
      var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - User Information").SWTObject("Label", "*").WndCaption;
      Log.Message(label );
       button.HoverMouse();
     ReportUtils.logStep_Screenshot("");
      button.Click();
    aqUtils.Delay(1000);
  }   
  
  
 ReportUtils.logStep_Screenshot(); 

 TextUtils.writeLog("Saving the changes"); 


var submitbtn =""
 
if((Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").Index==2)   &&
(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).isVisible()))
{
submitbtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
}
else if((Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").Index==2)   &&
(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).isVisible()))
{
submitbtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
}


     aqUtils.Delay(1000);
   


 Sys.HighlightObject(submitbtn);
     submitbtn.HoverMouse();
    ReportUtils.logStep_Screenshot(); 
    submitbtn.Click();
   aqUtils.Delay(1000);


 var approval = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
 //Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel2.TabControl
 //NameMapping.Sys.Maconomy.ObjectGroup.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");

  Sys.HighlightObject(approval);
    ReportUtils.logStep_Screenshot();
    Log.Message(approval.FullName);
    Sys.HighlightObject(approval);
    approval.Click();
    aqUtils.Delay(1000);
    ImageRepository.ImageSet.Maximize.Click();
   aqUtils.Delay(1000);
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    var All_approve = Aliases.ObjectGroup.AllApprovalActions;  
    All_approve.HoverMouse();
    ReportUtils.logStep_Screenshot(); 
    All_approve.Click();
   aqUtils.Delay(1000);
   Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
 var apprve_table = Aliases.ObjectGroup.ChangeUserApprovalTable;
 TextUtils.writeLog("Search For Approver"); 
  ReportUtils.logStep_Screenshot();

   var y=0;
  for(var x=0;x<apprve_table.getItemCount();x++){ 
      third_lvl_approver = true;
      var approvers="";
       if(apprve_table.getItem(x).getText_2(8)!="Approved"){
       approvers = apprve_table.getItem(x).getText_2(3).OleValue.toString().trim()+"*"+apprve_table.getItem(x).getText_2(4).OleValue.toString().trim();
       }
       
  
                       
    }
    
       getApproverDetails();
        ReportUtils.logStep_Screenshot();
       
var userInfoTab =
Aliases.ObjectGroup.UserApprovalClose;
userInfoTab.HoverMouse();
ReportUtils.logStep_Screenshot();
userInfoTab.Click();

var OpCo1 = EnvParams.Opco;
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco);
var OpCo2 = Approve_Level[0].replace(/OpCo -/g,OpCo1);
if((Approve_Level[0].indexOf(Project_manager)!=-1)||(OpCo2.indexOf(Project_manager)!=-1)){
  
level = 1;
var Approve = "";
Approve =
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(Approve)
Approve.HoverMouse();

ReportUtils.logStep_Screenshot();
Approve.Click();

ValidationUtils.verify(true,true,"Created User is Approved by" +Project_manager)
ReportUtils.logStep("INFO", "Created User is Approved by" +Project_manager);
TextUtils.writeLog("Approve User"); 

if(Approve_Level.length==1){
  aqUtils.Delay(1000);
  var Ok = 
  Sys.Process("Maconomy").SWTObject("Shell", "Users", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    Ok.HoverMouse();
    ReportUtils.logStep_Screenshot();
  Ok.Click();
   aqUtils.Delay(1000);
   
   
  if(Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
  var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
 
 expDate.HoverMouse();
    ReportUtils.logStep_Screenshot();
  expDate.Click();
  
  
  }
  }

  
    }
    }

      
      
}

function test()
{

    var savebtn =
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "")

    Sys.HighlightObject(savebtn);

   Delay(2000)

}


 function todo(lvl){ 
    var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  aqUtils.Delay(3000, Indicator.Text);
//  Delay(3000);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);

  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);

  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();

  
  
  aqUtils.Delay(15000, Indicator.Text);
//  Delay(15000);
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
if(lvl==0)
Client_Managt.DblClickItem("|Approve User Information (*)");

if(lvl==1)
Client_Managt.DblClickItem("|Approve User Information (Substitute) (*)");


break;
}
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


function aprvUser(ComId,EmpNo,userNmae,lvls,apvLvl){ 
aqUtils.Delay(5000, Indicator.Text);
if(ImageRepository.ImageSet.Show_Filter.Exists()){
aqUtils.Delay(2000, Indicator.Text);
ImageRepository.ImageSet.Show_Filter.Click();
aqUtils.Delay(2000, Indicator.Text);
} 

  
aqUtils.Delay(8000, Indicator.Text);
aqUtils.Delay(1000);
ReportUtils.logStep_Screenshot();

aqUtils.Delay(1000);
var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "")

firstcell.Keys(nameValue);
ReportUtils.logStep_Screenshot();
firstcell.Keys("[Tab][Tab]");
aqUtils.Delay(1000);
var Employee_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2)
Employee_no.Keys(employeeNo);
ReportUtils.logStep_Screenshot();
aqUtils.Delay(1000);
  
var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
  
var flag=false;
for(var v=0;v<table.getItemCount();v++){ 
if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(nameValue)){  
      flag=true;
      break;
    }else{ 
      table.Keys("[Down]");
    }
  } 
  
     ValidationUtils.verify(flag,true,"User is listed for Approval"); 
  if(flag){ 
var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
closefilter.Click();
ReportUtils.logStep_Screenshot();
aqUtils.Delay(3000);
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
var approve =Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8)
Sys.HighlightObject(approve);
if(approve.isEnabled()){ 
approve.HoverMouse();
ReportUtils.logStep_Screenshot();
  approve.Click();
ValidationUtils.verify(true,true,"User is Approved by "+userNmae)
 ReportUtils.logStep("INFO", "user is Approved by "+userNmae);
}
else{ 
  ReportUtils.logStep("INFO","Approve Button Is Invisible");
//  Log.Warning(ComId+" - "+JobNo +" - Approver :"+userNmae);
}
 aqUtils.Delay(1000);
  
  if(apvLvl==(ApproveInfo.length-1)){
if(lvls==0){
  


  var Ok = 
  Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
  Ok.HoverMouse();
    ReportUtils.logStep_Screenshot();
  Ok.Click();
   aqUtils.Delay(1000);
  if(Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
    
    var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
   expDate.HoverMouse();
  ReportUtils.logStep_Screenshot();
  expDate.Click();
  }
  
  }
 if(lvls==1){
   
 var Ok = 
  Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information (Substitute)", 1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    Ok.HoverMouse();
    ReportUtils.logStep_Screenshot();
  Ok.Click();
   Delay(2000);
  if(Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information (Substitute)").SWTObject("Composite", "", 2).SWTObject("Button", "OK").isVisible())
  {
var expDate = Sys.Process("Maconomy").SWTObject("Shell", "Approve User Information (Substitute)").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
    expDate.HoverMouse();
  ReportUtils.logStep_Screenshot(); 
  expDate.Click();
  }
   
   }
  
  }
  
  }
  

}

function getApproverDetails()
{
    var Approval_table = 
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
    linestatus = true;

    Sys.HighlightObject(Approval_table)
    for(var z=0;z<Approval_table.getItemCount();z++){ 
       approvers="";
       approvers = Approval_table.getItem(z).getText_2(3).OleValue.toString().trim()+"*"+Approval_table.getItem(z).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +z+ ": " +approvers);
       Approve_Level[y] = approvers;
       ReportUtils.logStep("INFO","Approver level :" +z+ ": " +Approve_Level[y]);
       y++;
    }

}

function CredentialLogin(){ 

for(var i=level;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  Log.Message(Approve_Level[i]);
  var Cred = Approve_Level[i].split("*");
  for(var j=0;j<2;j++){
    Log.Message(Cred[j])
Log.Message(j)
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     workBook = Project.Path+excelName;
     var sheetName = "Agency Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 
    workBook = Project.Path+excelName;
     Log.Message(workBook)
    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }
  else{ 
   var Eno =  Cred[j].substring(Cred[j].indexOf("(")+1,Cred[j].indexOf(")") )
    if(UserN){ 
      goToHR();
      UserN = false;
    }
    temp = searchNumber(Eno);
  }
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



function LogReport_name(ExcelData,value,JG){ 
var compStatus = "";
      for(var exl =0;exl<ExcelData.length;exl++){
        var splits = []; 
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
      if(splits[0]==value.toString().trim()){ 
        compStatus = ExcelData[exl]+"_"+JG;
        break;
      }
      }
Log.Message(compStatus);
return compStatus
}

function getExcelData_Company(rowidentifier,column) { 
excelData =[];  
//Log.Message(" ");
//Log.Message(excelName)
//Log.Message(workBook);
//Log.Message(sheetName);
var xlDriver = DDT.ExcelDriver(workBook,sheetName,false);
var id =0;
var colsList = [];
var temp ="";
//Log.Message(rowidentifier);
     while (!DDT.CurrentDriver.EOF()) {
//Log.Message(xlDriver.Value(0).toString().trim())
//Log.Message("Excel Column :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim().toUpperCase()==rowidentifier.toUpperCase()){
//       Log.Message("Row Identifier is Matched");
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     
     if(temp.indexOf("*")!=-1){
     var excelData =  temp.split("*");
//     Log.Message(excelData);
//     for(var i=0;i<comma_separator.length;i++){ 
//       
//     }
       
     }else if(temp.length>0){ 
      excelData[0] = temp;
//       excelData[0] = temp.substring(0, temp.indexOf("-"));
//       excelData[1] = temp.substring(temp.indexOf("-")+1)
     }
     
     DDT.CloseDriver(xlDriver.Name);
for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);
     return excelData;
  
}



 
 
