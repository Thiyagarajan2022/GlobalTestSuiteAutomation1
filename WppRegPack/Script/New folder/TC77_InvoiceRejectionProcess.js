﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "InvoiceRejection";
var invoiceDetails = [];
var approvers ="";
var Approve_Level = [];
var invoice = [];
var HRData = [];
var LoginEmp = [];
var y = 0;
var UserPasswd = [];
var UserPsd = [];
function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.AccountPayable.Exists()){
ImageRepository.ImageSet.AccountPayable.Click();// GL
}
else if(ImageRepository.ImageSet.AccountPayable2.Exists()){
ImageRepository.ImageSet.AccountPayable2.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
}

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|AP Transactions");
}

}
Delay(6000);
}



function ApprovalLevel(){ 
  invoice = SOXexcel(sheetName,1);
  var invoiceAllocations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  invoiceAllocations.Click();
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  firstcell.Click();
  Delay(2000);
  firstcell.Keys("[Tab]");
  var journalNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  journalNo.Click();
//  Delay(2000);
//  journalNo.setText(invoice[1]);
  journalNo.Keys("[Tab]");
  var transcationNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  transcationNo.Click();
  Delay(2000);
  transcationNo.setText(invoice[1]);
  transcationNo.Keys("[Tab][Tab][Tab][Tab]");
  var InvoiceNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  InvoiceNo.Click();
  Delay(2000);
  InvoiceNo.setText(invoice[2]);
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
      var itemCount = table.getItemCount();
      var flag = false;
      for(var i=0;i<itemCount;i++){
      if((table.getItem(i).getText_2(2).OleValue.toString().trim()==invoice[1])&&(table.getItem(i).getText_2(6).OleValue.toString().trim()==invoice[2])){ 
        flag = true;
          break;
      }
      else{ 
        table.Keys("[Down]");
      }
      }
if(flag) { 
if(ImageRepository.ImageSet.Close_Filter.Exists()){ 
  ImageRepository.ImageSet.Close_Filter.Click();
  Delay(5000);
}

var purchaseOrder = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
purchaseOrder.Click();
Delay(2000);
var AllApproved = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 9)
AllApproved.Click();
Delay(4000);
y =0 ;
var ApproverTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 8).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2)
var levels = [];
for(var i=0;i<ApproverTable.getItemCount();i++){ 
  
    approvers="";
       approvers = invoice[1]+"*"+invoice[2]+"*"+ApproverTable.getItem(i).getText_2(3).OleValue.toString().trim()+"*"+ApproverTable.getItem(i).getText_2(4).OleValue.toString().trim();
       Log.Message("Approver level :" +i+ ": " +approvers);
levels[y] = approvers;
       y++;
};
  
Approve_Level = Array.from(new Set(levels));
}
//HRData = WorkspaceUtils.goToHR();
LoginEmp = WorkspaceUtils.Credentiallogin(Project.Path+excelName, "userRoles");

//if(invoice[0]!="")
//UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,invoice[0]);
//else
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,"");

UserPsd = Array.from(new Set(UserPasswd));
RestMaconomy(UserPsd)

//  for(var i=0;i<Approve_Level.length;i++){ 
//    Log.Message("Data :"+Approve_Level[i]);
//  }
}



//function Login_Match(Approve_Level,LoginEmp,HRData,comID){ 
//login_satuts = true;
//Delay(3000);
//var UserPasswd = [];
//var z =0;
//for(var i=0;i<Approve_Level.length;i++){ 
//if((Approve_Level[i].indexOf("OpCo")!=-1) && (comID!="")){
//Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,comID);
//}
//// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
//if(Approve_Level[i].indexOf("SSC - Biller")==-1){
//Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
//}
//
//var tempLevel = Approve_Level[i].split("*");
//ifGotIT = true;
//var level = 0;
//for(var j=2;j<tempLevel.length;j++){ 
//level++;
//if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
//var temp = tempLevel[j].replace(" (","*");
//temp = temp.replace(")","");
////Log.Message("temp :"+temp)
//var tempSplit = temp.split("*");
//
//for(var k=0;k<LoginEmp.length;k++){
//  var A_temp = LoginEmp[k].split("*");
////    Log.Message("tempSplit[0] :"+tempSplit[0]);
////    Log.Message("A_temp[0] :"+A_temp[0]);
////    Log.Message("tempSplit[1] :"+tempSplit[1]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
// if(tempSplit[0]==A_temp[0]){ 
// if(level==1){
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
//    }
//    else{ 
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";  
//    }
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break; 
//      
// }else{ 
// if(tempSplit[1]==A_temp[2]){ 
// if(level==1){
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
//    }else{ 
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";
//    }
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }     
// }
//    
//}
//if(!ifGotIT){ 
//  break;
//}
//}
//  
//if((tempLevel[j].indexOf("SSC -")!=-1) || (tempLevel[j].indexOf("Central Team -")!=-1)){ 
////    Log.Message("tempLevel[j] :"+tempLevel[j]);
//   if(tempLevel[j].indexOf("Central Team - Client Management")!=-1){ 
//    temp2 = "Central Team - Client Account Management";
//  }
//  else if(tempLevel[j].indexOf("Central Team - Vendor Management")!=-1){ 
//    temp2 = "Central Team - Vendor Account Management";
//  }
//  else if(tempLevel[j].indexOf("SSC - Expense Cashiers")!=-1){ 
//    temp2 = "SSC - Cashier";
//  }else{ 
//    temp2 = tempLevel[j];
//  }
//for(var k=0;k<LoginEmp.length;k++){
//  var A_temp = LoginEmp[k].split("*");  
// if(temp2==A_temp[1]){ 
// if(level==1){
//   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
//   }
//   else{ 
//   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1"; 
//   }
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }
//}  
//
//if(!ifGotIT){ 
//  break;
//}
//}
//  
//  
//  
//if((tempLevel[j].indexOf(" (")==-1) && (tempLevel[j].indexOf(")")==-1) && 
//(tempLevel[j].indexOf("SSC -")==-1) && (tempLevel[j].indexOf("Central Team -")==-1)){ 
//    
//for(var k=0;k<LoginEmp.length;k++){
//  var A_temp = LoginEmp[k].split("*");
//  if(A_temp[0]==tempLevel[j]){  // Better  to use level[j].indexOf(LoginArrays[k])
//  if(level==1){
//  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
//  }else{ 
//  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";
//  }
//  Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }
//}
//if(!ifGotIT){ 
//  break;
//}
//}
//  
//if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
//
//var temp = tempLevel[j].replace(" (","*");
//temp = temp.replace(")","");
////Log.Message("temp :"+temp)
//var tempSplit = temp.split("*");
//
//for(var k=0;k<HRData.length;k++){
//  var A_temp = HRData[k].split("*");
////    Log.Message("tempSplit[0] :"+tempSplit[0]);
////    Log.Message("A_temp[0] :"+A_temp[0]);
////    Log.Message("tempSplit[1] :"+tempSplit[1]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
// if(tempSplit[1]==A_temp[1]){ 
// if(level==1){
//   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123"+"*"+"0";
//   }else{ 
//   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123"+"*"+"1";
//   }
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }
//    
//}
//if(!ifGotIT){ 
//  break;
//}
//}
// 
//  
//}
//if(ifGotIT){ 
//  Log.Warning("UserName and Password is Not Matched for Approver and Substitute :"+Approve_Level[i]);
//  login_satuts = false;
//  break;
//}
//  
//}
//
//return UserPasswd;
//}




function gotTODOs_Approve(level){ 
  var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.DBlClick();
  Delay(4000);
//  var refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
//  refresh.Click();
  
  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
  var refresh;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
if(refresh.isVisible()){ 
refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
refresh.Click();

  
  
  Delay(15000);
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
if(level==0)
Client_Managt.DblClickItem("|Approve Invoice Allocation Line (*)");

if(level==1)
Client_Managt.DblClickItem("|Approve Invoice Allocation Line (Substitute) (*)");


//var temp = Client_Managt.getItemCount();
//Log.Message(temp);
//for(var i=0;i<temp;i++){ 
//  if(Client_Managt.getItem(i).getText(0).OleValue.toString().trim().indexOf("Approve Invoice Allocation Line")!=-1){ 
//
//    
//  }
//  else{ 
//    Client_Managt.Keys("[Down]");
//    Delay(2000);
//  }
//}
break;
}
}
}

//Delay(8000);
//if(ImageRepository.ImageSet.Show_Filter.Exists()){ 
//  ImageRepository.ImageSet.Show_Filter.Click();
//  Delay(2000);
//}

}

function RestMaconomy(UserPasswd){ 
//var UserPasswd = [];
//UserPasswd[0] = "1702*Automation Client 19December2018 19:53:38*1707 - Finance*CORE@WPP123";;
//UserPasswd[1] = "122219*Regular Hindustan*somsubhra.banerjee@jwt.com*CORE@WPP123";
//UserPasswd[0] = "1706*Automation Client 19December2018 19:53:38*SSC IN -  CT Clients*CORE@WPP123";
Log.Message(UserPasswd.length);
for(var j=0;j<UserPasswd.length;j++){

var temp = UserPasswd[j];
var temp_user = temp.split("*");
var uname = temp_user[2]; 
Log.Message(uname)
var pwd = temp_user[3];
Log.Message(pwd)
WorkspaceUtils.Rests(uname,pwd);
   var ApproveFlag = true; 
gotTODOs_Approve(temp_user[4]);

//  var toDo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
//  toDo.DBlClick();
//  Delay(4000);
////  var refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
////  refresh.Click();
//  
//  var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
//  var refresh;
//Log.Message(childCC)
//for(var i=1;i<=childCC;i++){ 
//refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
//if(refresh.isVisible()){ 
//refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
//refresh.Click();
//
//  
//  
//  Delay(15000);
//Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
//if(Client_Managt.isVisible()){ 
//Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
////Client_Managt.DblClickItem("|Approve Invoice Allocation Line*");
//var temp = Client_Managt.getItemCount();
//Log.Message(temp);
//for(var i=0;i<temp;i++){ 
//if(ApproveFlag)
//  if(Client_Managt.getItem(i).getText(0).OleValue.toString().trim().indexOf("Approve Invoice Allocation Line")!=-1){ 







  

Delay(8000);
if(ImageRepository.ImageSet.Show_Filter.Exists()){ 
  ImageRepository.ImageSet.Show_Filter.Click();
  Delay(2000);
}
//var invoiceAllocations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
//  invoiceAllocations.Click();
  Delay(4000);
  var firstcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "");
//  firstcell.Click();
//  Delay(2000);
  firstcell.Keys("[Tab][Tab][Tab][Tab]");

  var InvoiceNumber = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  InvoiceNumber.Click();
  Delay(2000);
  InvoiceNumber.setText(temp_user[1]);
//  InvoiceNumber.setText("502");
  InvoiceNumber.Keys("[Tab][Tab]");
  var transcationNo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  transcationNo.Click();
  Delay(2000);
  transcationNo.setText(temp_user[0]);
//  transcationNo.Keys("1707100057");
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
      var itemCount = table.getItemCount();
      var flag = false;
      for(var i=0;i<itemCount;i++){
      if((table.getItem(i).getText_2(5).OleValue.toString().trim()==temp_user[1])&&(table.getItem(i).getText_2(7).OleValue.toString().trim()==temp_user[0])){ 
//      if((table.getItem(i).getText_2(5).OleValue.toString().trim()=="502")&&(table.getItem(i).getText_2(7).OleValue.toString().trim()=="1707100057")){
        flag = true;
        ApproveFlag = false;
          break;
      }
      }
if(flag) { 
if(ImageRepository.ImageSet.Close_Filter.Exists()){ 
  ImageRepository.ImageSet.Close_Filter.Click();
  Delay(5000);
}
var RejectAll = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 8);
Sys.HighlightObject(RejectAll)
RejectAll.Click();
Delay(6000);

var remarks = Sys.Process("Maconomy").SWTObject("Shell", "Reject").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
remarks.setText("Rejected");
Delay(2000);
var reject = Sys.Process("Maconomy").SWTObject("Shell", "Reject").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Reject all");
Sys.HighlightObject(reject)
reject.Click();
Delay(6000);
WorkspaceUtils.closeAllWorkspaces();

//var action = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("GroupToolItemControl", "", 9);
//action.Click();
//  Delay(3000);
//  Sys.Process("Maconomy").Refresh();
//  var table = Sys.Process("Maconomy").Window("#32768", "", 1);
//  Sys.HighlightObject(table);
//  Sys.Desktop.KeyDown(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyUp(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Sys.Desktop.KeyUp(0x28);
//  Delay(1000);
//  Sys.Desktop.KeyDown(0x0D);
//  Sys.Desktop.KeyUp(0x0D);
//  Delay(4000);


}
//------------
//}
//  else{ 
//    Client_Managt.Keys("[Down]");
//    Delay(2000);
//  }
//}
//break;
//}
//}
//}


//------------

}
}
function SOXexcel(CreateClient,start){ 
//function SOXexcel(){ 
 var Arrays = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, CreateClient, true);
//var xlDriver = DDT.ExcelDriver(Project.Path+excelName, "CreateClient", true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
//      for(var idx=1;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
//      }
     Arrays[id]=temp;
     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrays;
}



function VendorInvoiceRejection(){ 
  gotoMenu();
  ApprovalLevel();

  
}