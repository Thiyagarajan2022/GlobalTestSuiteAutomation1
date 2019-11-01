/*Closes workspaces after job completes in maconomy*/
function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}


function SearchByValue(ObjectAddrs,popupName,value){ 
var checkmark = false;
  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          checkmark = true;
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          Delay(1000);
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      Delay(1000);
      ObjectAddrs.setText("");
    }
    return checkmark;
}



function SearchByValues(ObjectAddrs,popupName,value){

var checkmark = false; 
  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    Delay(1000);
    var alljob = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
    alljob.Click();
    Delay(2000); 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", "Search ")
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
         OK.Click();
         checkmark = true;
         break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel")
           cancel.Click();
          Delay(1000);
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel")
      cancel.Click();
      Delay(1000);
      ObjectAddrs.setText("");
    }
    return checkmark;
}





function SearchByValuePicker(ObjectAddrs,popupName,value){ 
var checkmark =  false;
  Delay(1000);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    Delay(3000);
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    Delay(3000);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    Delay(5000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
          OK.Click();
          checkmark = true;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel")
          cancel.Click();
          Delay(1000);
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel")
      cancel.Click();
      Delay(1000);
      ObjectAddrs.setText("");
    }
    return checkmark;
}






function CalenderDateSelection(ObjectAddrs,value){ 
    var temp = "";
  temp = value.split("/");
  
  var leapYear = false;
  if((temp[2]>=1800) &&(temp[2]<2500)){
  if(temp[2]%4 == 0)
    {
        if( temp[2]%100 == 0)
        {
            // year is divisible by 400, hence the year is a leap year
            if ( temp[2]%400 == 0){ 
            leapYear = true;
            if((temp[0]==1)||(temp[0]==3)||(temp[0]==5)||(temp[0]==7)||(temp[0]==8)||(temp[0]==10)||(temp[0]==12)){ 
              if((temp[1]>0) &&(temp[1]<=31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
              
            }else{ 
            if(temp[0]==2){ 
              if((temp[1]>0) &&(temp[1]<30)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else if((temp[0]>0) &&(temp[0]<13)){ 
              if((temp[1]>0) &&(temp[1]<31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else{ 
               Log.Message("Month is invalid");
            }  
        }
    }      
        }
        
    }
     if(!leapYear){ 
        if((temp[0]==1)||(temp[0]==3)||(temp[0]==5)||(temp[0]==7)||(temp[0]==8)||(temp[0]==10)||(temp[0]==12)){ 
              if((temp[1]>0) &&(temp[1]<=31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
              
            }else{ 
            if(temp[0]==2){ 
              if((temp[1]>0) &&(temp[1]<29)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else if((temp[0]>0) &&(temp[0]<13)){ 
              if((temp[1]>0) &&(temp[1]<31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else{ 
               Log.Message("Month is invalid");
            }  
        }
  }
    }else{ 
      Log.Message("Year is invalid it should Between 1800-2499");
    }
}





function StartTime(){ 
    var dif;
    var STIME="";
    var TodayValue = aqDateTime.Today();
    var StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
    var EncodedDate = aqConvert.DateTimeToFormatStr(StringTodayValue,"%d%#B%Y"); 
//    Log.Message(EncodedDate)
    STIME = EncodedDate+" "+getFormattedCurrentTime();
//    Log.Message("Start DATE & TIME :"+EncodedDate +" "+STIME)
    var start = STIME.split(":");
    if(start[1]>0){ 
    dif = Number(start[2]) + Number(start[1]*60);
    }
    if(start[0]>0){ 
    dif = dif + Number(start[0]*60*60);
    }

return STIME;
}
function getFormattedCurrentTime(){
    TodayValue = aqConvert.DateTimeToFormatStr(aqDateTime.Time(), "%H:%M:%S");
    return TodayValue;
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
//for(var j=2;j<tempLevel.length;j++){ 
//
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
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }else{ 
// if(tempSplit[1]==A_temp[2]){ 
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
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
//   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
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
//  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]; 
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
//   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
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



function Login_Match(Approve_Level,LoginEmp,HRData,comID){ 
login_satuts = true;
Delay(3000);
var UserPasswd = [];
var z =0;
for(var i=0;i<Approve_Level.length;i++){ 
if((Approve_Level[i].indexOf("OpCo")!=-1) && (comID!="")){
Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,comID);
}
// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
if(Approve_Level[i].indexOf("SSC - Biller")==-1){
Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
}
if(Approve_Level[i].indexOf("SSC - Billers")!=-1){
Approve_Level[i] = Approve_Level[i].replace(/SSC - Billers/g,"SSC IN -  Biller");
}

var tempLevel = Approve_Level[i].split("*");
ifGotIT = true;
var level = 0;
for(var j=2;j<tempLevel.length;j++){ 
level++;
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
 if(level==1){
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
    }
    else{ 
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";  
    }
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
   Log.Message(UserPasswd[z]);
   z++;
   ifGotIT = false;
   break; 
      
 }else{ 
 if(tempSplit[1]==A_temp[2]){ 
 if(level==1){
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
    }else{ 
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";
    }
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
//    Log.Message("tempLevel[j] :"+tempLevel[j]);
   if(tempLevel[j].indexOf("Central Team - Client Management")!=-1){ 
    temp2 = "Central Team - Client Account Management";
  }
  else if(tempLevel[j].indexOf("Central Team - Vendor Management")!=-1){ 
    temp2 = "Central Team - Vendor Account Management";
  }
  else if(tempLevel[j].indexOf("SSC - Expense Cashiers")!=-1){ 
    temp2 = "SSC - Cashier";
  }else{ 
    temp2 = tempLevel[j];
  }
for(var k=0;k<LoginEmp.length;k++){
  var A_temp = LoginEmp[k].split("*");  
 if(temp2==A_temp[1]){ 
 if(level==1){
   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
   }
   else{ 
   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1"; 
   }
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
  if(level==1){
  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
  }else{ 
  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";
  }
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
 if(level==1){
   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123"+"*"+"0";
   }else{ 
   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123"+"*"+"1";
   }
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

return UserPasswd;
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
            Delay(5000);
            checkMark = true;
            break;
          }else{
//          Log.Message("i :"+i);
//          Log.Message(value+" "+value.length);
          
//        Log.Message(list.getItem(i).getText_2(0).OleValue.toString().trim()+" "+list.getItem(i).getText_2(0).OleValue.toString().trim().length); 
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





function goToHR(){ 
var HRData = [];
Delay(3000);
  closeAllWorkspaces();
  Delay(1000)
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
//if(ImageRepository.ImageSet.User1.Exists()){
//  ImageRepository.ImageSet.User1.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.User3.Exists()){
//  ImageRepository.ImageSet.User3.DblClick();// GL
//}
//else if(ImageRepository.ImageSet.User2.Exists()){
//  ImageRepository.ImageSet.User2.DblClick();// GL
//}


var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var HRitem;
Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(HRitem.isVisible()){ 
HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
HRitem.DblClickItem("|Users");
}

}

//var HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
//HRitem.DblClickItem("|Users");
Delay(5000);
//var ActiveUser = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Users");
//ActiveUser.Click();
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
return HRData;
}



function Credentiallogin(excelsheet,sheetname) {
var LoginEmp = [];
  var xlDriver = DDT.ExcelDriver(excelsheet, sheetname, false);
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
   return LoginEmp;
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
    
  /*  
    Delay(7000);
    var status = true;
    while(status){
    var Name = Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption.toString().trim();
    if(Name=="Login to Deltek Maconomy"){ 
      status = false;
    usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
    pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");

      Delay(2000);
//      Login_to_Deltek_Maconomy();
      usernameAddr.SetFocus();
      usernameAddr.setText(loginuser);
      pwdAddr.setText(loginpassword);
      btnLogin.click();
      Delay(10000);
      break;
    }
    if(Name=="Server Configuration"){ 
      Delay(2000);
    server_link = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Text", "");
    port_number = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2)
    company_name = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 3)
    chk_box = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Button", "Do not ask me again");
    connectbtn = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Connect");

//      Server_Configuration();
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
    Delay(10000); 
    
    */  
}
