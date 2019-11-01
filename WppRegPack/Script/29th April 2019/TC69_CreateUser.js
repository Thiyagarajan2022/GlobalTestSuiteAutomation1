//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var userInfo = [];
var Approve_Level = [];
var UserPasswd = [];
var sheetName = "UserCreation";
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
function goToMenu(){ 
  
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
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


if(ImageRepository.ImageSet.User1.Exists()){
  ImageRepository.ImageSet.User1.DblClick();// GL
}
else if(ImageRepository.ImageSet.User3.Exists()){
  ImageRepository.ImageSet.User3.DblClick();// GL
}
else if(ImageRepository.ImageSet.User2.Exists()){
  ImageRepository.ImageSet.User2.DblClick();// GL
}
Delay(3000);
Log.Message("User Icon is Clicked");
}


function goToUsers(){ 
Delay(3000)
 var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
 All_User.Click();
//  var employees = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
//  employees.DblClickItem("|Employees");
  var Add_Visible0 = true;
  var New_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("SingleToolItemControl", "", 3,60000);
  while(Add_Visible0){
  if(New_User.isEnabled()){
  New_User.Click();
  Add_Visible0 = false;
  }
  }
  
}


function UserInformation(){ 
//userInfo =  excel(sheetName);
Log.Message("User Creation is Started");

  Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).WaitSWTObject("McTextWidget", "", 2,60000);
  Delay(1000);
  var name = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
  if(userInfo[0]!=""){
  name.Keys("^a[BS]")
  name.setText(userInfo[0]+" "+STIME);
  }else{ 
    ValidationUtils.verify(false,true,"User Name is Needed to Create a User");
  }
  
var employee = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2)
 if(userInfo[1]!=""){
   employee.Click();
   WorkspaceUtils.SearchByValue(employee,"Employee",userInfo[1]);
   } 
    
var company = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3)
 if(userInfo[2]!=""){
   company.Click();
   WorkspaceUtils.SearchByValue(company,"Company",userInfo[2]);
   }
 
var user_type = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 2).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 2);
 if(userInfo[3]!=""){
   user_type.Click();
   WorkspaceUtils.SearchByValue(user_type,"User Type",userInfo[3]);
   }
 
  
  
  var periodfrom = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 2)
  if(userInfo[4]!=""){
  WorkspaceUtils.CalenderDateSelection(periodfrom,userInfo[4])
  }


  var periodto = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "").SWTObject("McDatePickerWidget", "", 4)
  if(userInfo[5]!=""){
  WorkspaceUtils.CalenderDateSelection(periodto,userInfo[5])
  }  

  var template = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "", 4).SWTObject("Composite", "").SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  if((template.getSelection()) && ((userInfo[6]=="No")||(userInfo[6]=="no"))) { 
    template.Click();
      Log.Message("Templete is UnChecked")
    }
  if((!template.getSelection()) && ((userInfo[6]=="Yes")||(userInfo[6]=="Yes"))) { 
    template.Click();
      Log.Message("Templete is Checked")
    }
  
  Delay(2000);
  var create = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");
  Sys.HighlightObject(create);
  create.Click();
  Delay(8000);
//  var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
//  Sys.HighlightObject(cancel);

var pop = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users");

if(pop.Child(1).JavaFullClassName.indexOf('Label')!=-1){
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Composite", "", 2).SWTObject("Button", "OK");
var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",1).SWTObject("Label", "*");
Log.Message(label.getText());
//button.Click();
ImageRepository.ImageSet.OK_Button.Click();

  Delay(2000);
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",2).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",2).SWTObject("Label", "*");
Log.Message(label.getText());
//button.Click();
ImageRepository.ImageSet.Ok.Click();

  Delay(2000);
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",3).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",3).SWTObject("Label", "*");

Log.Message(label.getText());
//button.Click();
ImageRepository.ImageSet.Ok.Click();

  Delay(2000);
//Sys.Process("Maconomy").SWTObject("Shell", "Users - Users", 4).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
if(ImageRepository.ImageSet.Ok.Exists()){
var button = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",4).SWTObject("Composite", "", 2).SWTObject("Button", "OK")
var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users",4).SWTObject("Label", "*");
Log.Message(label.getText());
//button.Click();
ImageRepository.ImageSet.Ok.Click();
  Delay(2000);
  }
 
Log.Message("User is Created");
return true; 
  }else{ 
    var label = Sys.Process("Maconomy").SWTObject("Shell", "Users - Users").SWTObject("Text", "");
    Log.Message(label.getText());
    ImageRepository.ImageSet.OK_Button.Click();
    Delay(3000);
    
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create User").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
    Sys.HighlightObject(cancel);
    cancel.Click();
    
    ValidationUtils.verify(false,true,"User is Not Created");
    return false;
  }
   
   

}
 
 
 
 
 
 
 
 
 
  
  
  
function user_Approval(){ 
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  firstcell.Keys(userInfo[0]+" "+STIME);
//  firstcell.Keys("Muthu 18:42:32");
  firstcell.Keys("[Tab][Tab]");
  Delay(2000);
  var Employee_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
//  Employee_no.Keys("170710134");
  Employee_no.Keys(userInfo[1]);

  
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Delay(4000);
  
    var flag=false;
    for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userInfo[0]+" "+STIME)){ 
//      if(table.getItem(v).getText_2(0).OleValue.toString().trim()=="Muthu 18:42:32"){ 
        flag=true;
        Log.Message("User is created");
        break;
      }else{ 
        table.Keys("[Down]");
      }
    }
    
    
    
    
//    var flag = table.getItemCount()>0;
    ValidationUtils.verify(flag,true,"User Created is available in system");
  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    closefilter.Click();
    Delay(3000);
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
    var Access_Level = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
    Sys.HighlightObject(Access_Level);
    Access_Level.Click();
    Delay(3000);
    
    var Acc_lvl = Access_lvl_excel();
    for(var z=0;z<Acc_lvl.length;z++){
    var Add_Click = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
    Sys.HighlightObject(Add_Click);
    Add_Click.Click();
    Delay(3000);
//    if(z==0){
    var Acc_Code = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    }else{
//    var Acc_Code = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    }
    Sys.HighlightObject(Acc_Code);
    Acc_Code.Click();
//    Log.Message(Acc_lvl[z]);
    
    
  /*    
    
    Acc_Code.Keys("^a[BS]");
    Acc_Code.Keys(Acc_lvl[z]);
    Delay(4000);
    
    
    
  
    var op = Sys.Process("Maconomy").SWTObject("Shell", "");
    var sea = "Sys.Process(Maconomy).SWTObject(Shell, ).SWTObject(McValuePickerFooter, , 1)";
    var sea1 = "Sys.Process(Maconomy).SWTObject(Shell, ).SWTObject(McValuePickerFooter, , 1).SWTObject(CLabel, , 1)";
    var status = "fail";
    for(var i=0;i<op.ChildCount;i++){ 
    var test = op.Child(i).FullName.toString().trim();

    test = test.replace(/"/g,"");
    if(test==sea){ 
    var tt= op.Child(i).FullName;

    for(var j=0;j<op.Child(i).ChildCount;j++){ 
      var test1 = op.Child(i).Child(j).FullName;
       Log.Message(test1);
        test1 = test1.replace(/"/g,"");
      if(test1==sea1){ 
    status = "pass";
    }
}
}
}
    if(status == "pass"){
    var search = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("McValuePickerFooter", "", 1).SWTObject("CLabel", "", 1)
    Sys.HighlightObject(search);
    search.Click();
//    var table =
    Delay(2000);
*/

    Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//    Log.Message("z :"+z +"  Acc_lvl[z]"+Acc_lvl[z]);
    code.setText(Acc_lvl[z]);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
//    Log.Message(table.getItemCount());
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Acc_lvl[z]){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          Delay(2000);
//       if(z==0){
       var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//       }else{
//       var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
//       }
       save.Click();
       Delay(5000);
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Access Level").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          Acc_Code.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Option").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
      Acc_Code.setText("");
    } 
/*
    }
     if(status == "fail"){
      Acc_Code.setText("");
      Log.Warning("This value is not available in Maconomy for Access Level");
    }
  */  
    }

//   var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
   var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 7);
    submit.Click();
    Delay(5000);
    var approval = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
    approval.Click();
    Delay(3000);
    ImageRepository.ImageSet.Maximize.Click();
    Delay(2000);
    var All_approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
    All_approve.Click();
    
//   var apprve_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
   var apprve_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);;
   var y=0;
  for(var x=0;x<apprve_table.getItemCount();x++){ 
      third_lvl_approver = true;
      var approvers="";
       if(apprve_table.getItem(x).getText_2(8)!="Approved"){
       approvers = apprve_table.getItem(x).getText_2(3).OleValue.toString().trim()+"*"+apprve_table.getItem(x).getText_2(4).OleValue.toString().trim();
//       Log.Message("Approver level :" +z+ ": " +approvers);
Log.Message("User Approver level : " +x+ " Approver :" +approvers);
       if(userInfo[1]!=""){
       Approve_Level[y] = userInfo[0]+" "+STIME+"*"+userInfo[1]+"*"+approvers;
       }else{ 
       Approve_Level[y] = userInfo[0]+" "+STIME+"**"+approvers;  
       }
       Log.Message(Approve_Level[y])
       y++;
       }
    }
   
  }
  
  
}

function Login_Match(){ 
login_satuts = true;
Delay(3000);
UserPasswd = [];

goToHR();

Credentiallogin();
var z =0;
for(var i=0;i<Approve_Level.length;i++){ 
if((Approve_Level[i].indexOf("OpCo")!=-1) && (userInfo[2]!="")){
Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,userInfo[2]);
}
// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
  if(Approve_Level[i].indexOf("SSC - Biller")==-1){
  Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
  }

var tempLevel = Approve_Level[i].split("*");
ifGotIT = true;
for(var j=2;j<tempLevel.length;j++){ 

if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
//Log.Message("temp :"+temp)
var tempSplit = temp.split("*");

  for(var k=0;k<LoginEmp.length;k++){
    var A_temp = LoginEmp[k].split("*");
//    Log.Message("tempSplit[0] :"+tempSplit[0]);
//    Log.Message("A_temp[0] :"+A_temp[0]);
//    Log.Message("tempSplit[1] :"+tempSplit[1]);
//    Log.Message("A_temp[1] :"+A_temp[1]);
   if(tempSplit[0]==A_temp[0]){ 
      UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
     Log.Message(UserPasswd[z]);
     z++;
     ifGotIT = false;
     break;     
   }else{ 
   if(tempSplit[1]==A_temp[2]){ 
      UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
     Log.Message(UserPasswd[z]);
     z++;
     ifGotIT = false;
     break;     
   }     
   }
    
  }
  if(!ifGotIT){ 
    break;
  }
  }
  
 if((tempLevel[j].indexOf("SSC -")!=-1) || (tempLevel[j].indexOf("Central Team -")!=-1)){ 
    
     if(tempLevel[j].indexOf("Central Team - Client Management")!=-1){ 
      temp2 = "Central Team - Client Account Management";
    }
    else if(tempLevel[j].indexOf("Central Team - Vendor Management")!=-1){ 
      temp2 = "Central Team - Vendor Account Management";
    }
    else if(tempLevel[j].indexOf("SSC - Expense Cashiers")!=-1){ 
      temp2 = "SSC - Cashier";
    }
  for(var k=0;k<LoginEmp.length;k++){
    var A_temp = LoginEmp[k].split("*");  
   if(tempLevel[j]==A_temp[1]){ 
     UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
     
     Log.Message(UserPasswd[z]);
     z++;
     ifGotIT = false;
     break;     
   }
  }  

  if(!ifGotIT){ 
    break;
  }
  }
  
  
  
 if((tempLevel[j].indexOf(" (")==-1) && (tempLevel[j].indexOf(")")==-1) && 
  (tempLevel[j].indexOf("SSC -")==-1) && (tempLevel[j].indexOf("Central Team -")==-1)){ 
    
  for(var k=0;k<LoginEmp.length;k++){
    var A_temp = LoginEmp[k].split("*");
    if(A_temp[0]==tempLevel[j]){  // Better  to use level[j].indexOf(LoginArrays[k])
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]; 
    Log.Message(UserPasswd[z]);
     z++;
     ifGotIT = false;
     break;     
   }
}
  if(!ifGotIT){ 
    break;
  }
  }
  
if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){

var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
//Log.Message("temp :"+temp)
var tempSplit = temp.split("*");

  for(var k=0;k<HRData.length;k++){
    var A_temp = HRData[k].split("*");
//    Log.Message("tempSplit[0] :"+tempSplit[0]);
//    Log.Message("A_temp[0] :"+A_temp[0]);
//    Log.Message("tempSplit[1] :"+tempSplit[1]);
//    Log.Message("A_temp[1] :"+A_temp[1]);
   if(tempSplit[1]==A_temp[1]){ 
     UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
     Log.Message(UserPasswd[z]);
     z++;
     ifGotIT = false;
     break;     
   }
    
  }
  if(!ifGotIT){ 
    break;
  }
  }
 
  
  }
  if(ifGotIT){ 
    Log.Warning("UserName and Password is Not Matched for Approver and Substitute :"+Approve_Level[i]);
    login_satuts = false;
    break;
  }
  
}

return login_satuts;
}


function login() {
    var xlDriver = DDT.ExcelDriver(workBook, sscCredential, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
//      Log.Message(temp)
     LoginArrays[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}


function Credentiallogin() {
    var xlDriver = DDT.ExcelDriver(Project.Path+ExcelFileName, Credential, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
//      Log.Message(temp)
     LoginEmp[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}

function logins() {
    var xlDriver = DDT.ExcelDriver(workBook, loginpassword, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
//      Log.Message(temp)
     LoginArr[id]=temp;
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
}


function goToHR(){ 
  Delay(3000);
    closeAllWorkspaces();


 if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();  
}

var HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
HRitem.DblClickItem("|Users");
Delay(5000);
var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
All_User.Click();
Delay(5000);
var HRTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
var z=0;
for(var i=0;i<HRTable.getItemCount();i++){ 
  if(HRTable.getItem(i).getText(2)!=""){
HRData[z] = HRTable.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+HRTable.getItem(i).getText_2(2).OleValue.toString().trim()
//Log.Message(HRData[z]);
z++;

}
}

  }

  

function gotoApprove(){
Delay(3000)
 var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
 All_User.Click(); 
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  firstcell.Keys(userInfo[0]+" "+STIME);
//  firstcell.Keys("Muthu");
  firstcell.Keys("[Tab][Tab]");
  Delay(2000);
  var Employee_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 2);
//  Employee_no.Keys("170710134");
  Employee_no.Keys(userInfo[1]);
  Delay(4000);
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(0).OleValue.toString().trim()==(userInfo[0]+" "+STIME)){ 
      flag=true;
      break;
    }else{ 
      table.Keys("[Down]");
    }
  }
  
  
  
  if(flag){ 
   var closefilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
    closefilter.Click();
    Delay(3000);
    Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).Refresh();
  var approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
  Sys.HighlightObject( approve );

  approve.Click();
  Delay(5000);
  var Ok = Sys.Process("Maconomy").SWTObject("Shell", "Users").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
  Ok.Click();
   WorkspaceUtils.closeAllWorkspaces();
}
}




function excel(sheetName){ 
 var Arrays = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=1;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim();
      }
      else{ 
        temp = temp;
      }
      }
     Arrays[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}


function Access_lvl_excel(){ 
 var Arrays = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=0;idx<colsList.length-1;idx++){  
       if((xlDriver.Value(colsList[idx])!=null)&& (xlDriver.Value(colsList[idx]).toString().trim()=="Access Level")){
          Arrays[id]=xlDriver.Value(colsList[idx+1]).toString().trim();
          id++;
      }
         }
     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}



function RestMaconomy(){ 
for(var i=0;i<UserPasswd.length;i++){

var temp = UserPasswd[i];
temp_user = [];
temp_user = temp.split("*");
var uname = temp_user[2]; 
//Log.Message(uname)
var pwd = temp_user[3];
//Log.Message(pwd)
Rests(uname,pwd);
    
goToMenu();
gotoApprove();
ValidationUtils.verify(true,true,"User has Approved by "+uname);
}
}

function Rests(uname,pwd){ 
Delay(5000);
      Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x52); //R 
     Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
     Sys.Desktop.KeyUp(0x52); //R
Delay(65000);
     var usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
    var pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    var btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
    usernameAddr.SetFocus();
    usernameAddr.setText(uname);
    pwdAddr.setText(pwd);
    btnLogin.click();
    Delay(10000);   
}

function CreateUser(){ 
STIME = WorkspaceUtils.StartTime();
  goToMenu();
  goToUsers();
  userInfo =  excel(sheetName);
  var status = UserInformation();
//var status = true;
  if(status){
  user_Approval();
//Approve_Level[0] = " Muthu 22:06:31*170710134*SACHINDRA P KARKERA (170710011)*";
  status = Login_Match();
  
  if(status){
  RestMaconomy()
//  gotoApprove()
  }
  }
  WorkspaceUtils.closeAllWorkspaces();
}




