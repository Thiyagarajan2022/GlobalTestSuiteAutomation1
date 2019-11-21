﻿//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
var sheetName = "Server Details";
  Indicator.Show();
  Indicator.PushText("waiing for window to open");
ExcelUtils.setExcelName(workBook, sheetName, true);
var usernameAddr;
   var pwdAddr;
   var dropdown;
   var btnLogin;
   var server_link;
   var port_number;
   var company_name;
   var chk_box;
   var connectbtn;
   var loginuser="";
   var loginpassword="";
   var server="";
   var port="";
   var company="";
   var loginName = "*";
   var serverName ="*";
   


function menu_link(){
    var Obj_menu= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    return Obj_menu;
    }
    
function login() {
  
Sys.Refresh();
var sysCount = Sys.ChildCount;
var process = false;
for(var cc=0;cc<sysCount;cc++){
if(Sys.Child(cc).ProcessName=="Maconomy")
process = true
}
  

if(process){
}

loginuser = "";
loginpassword = "";
var colsList = [];
var login_colsList = [];
email_ID =[];

TestedApps.Maconomy.Run();
aqUtils.Delay(20000, "Waiting for Maconomy to Start");
var status = true;
while(status){
var mainparent = Sys.Process("Maconomy")
aqUtils.Delay(5000, "Waiting to find Object");
Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy");
aqUtils.Delay(5000, "Waiting to find Child branch");
var childCount = Sys.Process("Maconomy").ChildCount;
for(var ci=0;ci<childCount;ci++){ 
  if((mainparent.Child(ci).Name!="JavaRuntime()")&&(mainparent.Child(ci).Visible!=false)){
  Full_Name = mainparent.Child(ci).WndCaption.toString().trim();
if((Full_Name.indexOf("Login to Deltek Maconomy")!=-1) ||(Full_Name.indexOf("Server Configuration")!=-1) ||
(Full_Name.indexOf("Inicio de sesión en Deltek Maconomy")!=-1) ||(Full_Name.indexOf("Configuración de servidor")!=-1) ||
(Full_Name.indexOf("登录到 Deltek Maconomy")!=-1) ||(Full_Name.indexOf("Server Configuration")!=-1)){
var Name = Full_Name;

status = false;
if((Name=="Login to Deltek Maconomy")|| (Name=="Inicio de sesión en Deltek Maconomy")||(Name=="登录到 Deltek Maconomy")){
if(Name=="Login to Deltek Maconomy"){ 
loginName = "Login to Deltek Maconomy";
serverName = "Server Configuration";
}else if(Name=="Inicio de sesión en Deltek Maconomy"){ 
loginName = "Inicio de sesión en Deltek Maconomy";
serverName = "Configuración de servidor";
}else{ 
loginName = "登录到 Deltek Maconomy";
serverName = "Server Configuration";
}
LoginMaconomy();
}
if((Name=="Server Configuration")||(Name=="Configuración de servidor")||(Name=="Server Configuration")){ 
aqUtils.Delay(2000, Indicator.Text);
  if(Name=="Server Configuration"){ 
    loginName = "* Deltek Maconomy";
    serverName = "Server Configuration";
  }else if(Name=="Configuración de servidor"){ 
    loginName = "Inicio de sesión en Deltek Maconomy";
    serverName = "Configuración de servidor";
  }else{ 
    loginName = "登录到 Deltek Maconomy";
    serverName = "Server Configuration";
  }
ServerConfigration();

}
break;
}
    
}
}
}
    
}
   
   function LoginAddress(){
    usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", loginName).SWTObject("Composite", "").WaitSWTObject("Composite", "", 1,60000).SWTObject("Text", "", 1);
    pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", loginName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    dropdown = Sys.Process("Maconomy").SWTObject("Shell", loginName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Combo", "");
    btnLogin = Sys.Process("Maconomy").SWTObject("Shell", loginName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
    }
    
    function ServerAddress(){
    server_link = Sys.Process("Maconomy").SWTObject("Shell", serverName).SWTObject("Composite", "").WaitSWTObject("Composite", "",1,60000).SWTObject("Text", "");
    port_number = Sys.Process("Maconomy").SWTObject("Shell", serverName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2)
    company_name = Sys.Process("Maconomy").SWTObject("Shell", serverName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 3)
    chk_box = Sys.Process("Maconomy").SWTObject("Shell", serverName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Button", "Do not ask me again");
    connectbtn = Sys.Process("Maconomy").SWTObject("Shell", serverName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Connect");
    }
    
    
function LoginMaconomy(){ 

var sheetName = "Server Details";

ExcelUtils.setExcelName(workBook, sheetName, true);
loginuser = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if((loginuser==null)||(loginuser=="")){ 
ValidationUtils.verify(false,true,"UserName is Needed to Login Maconomy");
}
loginpassword = ExcelUtils.getRowDatas("password",EnvParams.Opco)
if((loginpassword==null)||(loginpassword=="")){ 
ValidationUtils.verify(false,true,"Login Password is Needed to Login Maconomy");
}

Language = EnvParams.Language;

var sheetName = "LanguageLookUpTable";
ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"), sheetName, true);
Language = ExcelUtils.getRowDatas(Language,"MaconomyValue")
if((Language==null)||(Language=="")){ 
ValidationUtils.verify(false,true,"Language is Needed to Login Maconomy");
}
      aqUtils.Delay(2000, Indicator.Text);
      LoginAddress();
      aqUtils.Delay(1000, Indicator.Text);
      usernameAddr.SetFocus();
      usernameAddr.setText(loginuser);
      pwdAddr.setText(loginpassword);
      dropdown.DropDown();
      dropdown.ClickItem(Language);
      btnLogin.click();
       
    }
    
    
function ServerConfigration(){ 
var workBook = Project.Path+TextUtils.GetProjectValue("EnvDetailsPath")
var sheetName = "ServerDetails";

ExcelUtils.setExcelName(workBook, sheetName, true);
server = ExcelUtils.getRowDatas("serverAddress",EnvParams.instanceData )
if((server==null)||(server=="")){ 
ValidationUtils.verify(false,true,"Server Address is Needed to Login Maconomy");
}
port = ExcelUtils.getRowDatas("port",EnvParams.instanceData)
if((port==null)||(port=="")){ 
ValidationUtils.verify(false,true,"Port Number is Needed to Login Maconomy");
}
company = ExcelUtils.getRowDatas("company",EnvParams.instanceData)
if((company==null)||(company=="")){ 
ValidationUtils.verify(false,true,"Company is Needed to Login Maconomy");
}
aqUtils.Delay(2000, Indicator.Text);
ServerAddress();
server_link.SetFocus();
server_link.setText(server);
port_number.SetFocus();
port_number.setText(port);
company_name.SetFocus();
company_name.setText(company);
    
if(!chk_box.getSelection()){
chk_box.ClickButton(cbChecked);
}
connectbtn.click();
LoginMaconomy();
}
 