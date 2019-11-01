//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT PdfUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "Server Details";

var usernameAddr;
   var pwdAddr;
   var btnLogin;
   var server_link;
   var port_number;
   var company_name;
   var chk_box;
   var connectbtn;
   var loginuser="";
   var loginpassword="";
   var server="https://wpp-bautestapac.deltekenterprise.com";
   var port="443";
   var company="bauapac";
   


function menu_link(){
    var Obj_menu= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    return Obj_menu;
    }
    
function login(username,password) {
Delay(5000);
      Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //R 
     Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
     Sys.Desktop.KeyUp(0x58);

   loginuser = username;
   loginpassword = password;
   var colsList = [];
   var login_colsList = [];
   email_ID =[];
  
//   var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
//    var i =0;
//     for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){ 
//       login_colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
//     }
//       server_link = xlDriver.Value(login_colsList[1]).toString().trim() ;
//       Log.Message(server_link);
//       xlDriver.Next();
//       port_number = xlDriver.Value(login_colsList[1]).toString().trim() ;
//       Log.Message(port_number);
//       xlDriver.Next();
//       company_name = xlDriver.Value(login_colsList[1]).toString().trim() ;
//       Log.Message(company_name);
//       xlDriver.Next();
//       loginuser = xlDriver.Value(login_colsList[1]).toString().trim() ;
//       Log.Message(loginuser);
//       xlDriver.Next();
//       loginpassword = xlDriver.Value(login_colsList[1]).toString().trim() ;
//       Log.Message(loginpassword);
//       xlDriver.Next();
//    
//     DDT.CloseDriver(xlDriver.Name);


    TestedApps.Maconomy.Run();
    Delay(7000);
        var status = true;
    while(status){
    var Name = Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption.toString().trim();
    if(Name=="Login to Deltek Maconomy"){ 
      status = false;
      Delay(2000);
      Login_to_Deltek_Maconomy();
      usernameAddr.SetFocus();
      usernameAddr.setText(loginuser);
      pwdAddr.setText(loginpassword);
      btnLogin.click();
      Delay(10000);
      break;
    }
    if(Name=="Server Configuration"){ 
      Delay(2000);
      Server_Configuration();
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
      Delay(5000);
    }
    }
    
   }
   
   function Login_to_Deltek_Maconomy(){
    usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
    pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
    }
    
    function Server_Configuration(){
    server_link = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Text", "");
    port_number = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2)
    company_name = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 3)
    chk_box = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Button", "Do not ask me again");
    connectbtn = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Connect");
    }
    
 