//USEUNIT ExcelUtils
//USEUNIT TextUtils
var instanceData= "";
var businessFlow = null;
var Opco = null;
var TestingType = "";
var Country = "";
var CountryList = "";
var Language = "";
var OpcoNumber = "";
var path = "";
var testcase = "";
var Lang = "";
var OpcoNum = null;
var Lang_Jenk = null;
var Vname = null;
var Pname = null;
var Cname = null;
function getEnvironment(){
var i;
var nArgs = BuiltIn.ParamCount();
var result = null;
var sheetLoc = null;
var colsList=[];


var xlDriver= DDT.ExcelDriver(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"EnvDetails",true);

//instanceData = "BAUTESTAPAC"
//Country = "India"
//testcase = "CreatePurchseOrder";
//TestingType = "SysTest"
//OpcoNum = 1707;
//Lang_Jenk = "English";


for (i = 1; i <= nArgs ; i++){    
//Log.Message(BuiltIn.ParamStr(i));
if(BuiltIn.ParamStr(i).indexOf("Environment")!=-1){
   var inst = BuiltIn.ParamStr(i);
   instanceData = (inst.substring(inst.indexOf(":"))).trim();
   if(instanceData!=null)
   instanceData = instanceData.substring(instanceData.indexOf("=")+2,instanceData.length-1);
   Log.Message("Environment :"+instanceData);     

}

if(BuiltIn.ParamStr(i).indexOf("Country")!=-1){
   var inst = BuiltIn.ParamStr(i);
   CountryList = (inst.substring(inst.indexOf(":"))).trim(); 
//   if(Country!=null)
//   Country = Country.substring(Country.indexOf("=")+2,Country.length-1);     
//   Log.Message("Country :"+Country); 
   
   if(CountryList!=null){
   CountryList = CountryList.substring(CountryList.indexOf("=")+2,CountryList.length-1);
   if(CountryList!="ALL"){
   if(CountryList.indexOf(",")!=-1){
   var temp = CountryList.split(",");
   Country = temp[0];
   }
   else{ 
    Country =  CountryList;
   }
   }
   }
   Log.Message("CountryList :"+CountryList);
//   Log.Message("Country :"+Country);
}



if(BuiltIn.ParamStr(i).indexOf("TestCases")!=-1){
   var inst = BuiltIn.ParamStr(i);
   testcase = (inst.substring(inst.indexOf(":"))).trim(); 
   if(testcase!=null)
   testcase = testcase.substring(testcase.indexOf("=")+2,testcase.length-1);
   Log.Message("testcase :"+testcase);     

}

if(BuiltIn.ParamStr(i).indexOf("TestingType")!=-1){
   var inst = BuiltIn.ParamStr(i);
   TestingType = (inst.substring(inst.indexOf(":"))).trim(); 
   if(TestingType!=null)
   TestingType = TestingType.substring(TestingType.indexOf("=")+2,TestingType.length-1);
   Log.Message("TestingType :"+TestingType);     

}

if(BuiltIn.ParamStr(i).indexOf("OpCo")!=-1){
   var inst = BuiltIn.ParamStr(i);
  OpcoNum = (inst.substring(inst.indexOf(":"))).trim(); 
   if(OpcoNum!=null){
   OpcoNum = OpcoNum.substring(OpcoNum.indexOf("=")+2,OpcoNum.length-1);
   if(OpcoNum!="ALL"){
   if(OpcoNum.indexOf(",")!=-1){
   var temp = OpcoNum.split(",");
   Opco = temp[0];
   }
   else{ 
    Opco =  OpcoNum;
   }
   }
   }
   Log.Message("Opco :"+OpcoNum);     

}

if(BuiltIn.ParamStr(i).indexOf("Language")!=-1){
   var inst = BuiltIn.ParamStr(i);
  Lang_Jenk = (inst.substring(inst.indexOf(":"))).trim(); 
   if(Lang_Jenk!=null){
   Lang_Jenk = Lang_Jenk.substring(Lang_Jenk.indexOf("=")+2,Lang_Jenk.length-1);
   Language = Lang_Jenk;
   Log.Message(Language);
   }

//   Log.Message("Opco :"+Opco);     
//Language = "English";

}

}

//instanceData = "TESTAPAC"
//Country = "China"
//testcase = "CreatePurchaseOrder";
//TestingType = "SysTest"
//OpcoNum = "1307";
//Language = "English";
//Opco = "1307"

//if(result==null){
//  result =  "InstanceToRun";
//}
for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
 colsList[idx] = DDT.CurrentDriver.ColumnName(idx);     
}
 while (!DDT.CurrentDriver.EOF()) { 
   
  for(var idx=0;idx<colsList.length-1;idx++){
    
  if((instanceData==null)||(instanceData=="")){
  if(("InstanceToRun".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  instanceData = xlDriver.Value(colsList[idx+1]).toString().trim();
  }
  }
  
  if((TestingType==null)||(TestingType=="")){
  if(("Testing Type".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  TestingType = xlDriver.Value(colsList[idx+1]).toString().trim();    
  
  }
  }
  
  if((CountryList==null)||(CountryList=="")){
  if(("Country".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  Country = xlDriver.Value(colsList[idx+1]).toString().trim();    

  }
  }
  
  if((OpcoNum==null)||(OpcoNum=="")){
  if(("Opco".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  Opco = xlDriver.Value(colsList[idx+1]).toString().trim();    

  }
  }
  
  if((Lang_Jenk==null)||(Lang_Jenk=="")){
  if(("Language".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  Language = xlDriver.Value(colsList[idx+1]).toString().trim();    

  }
  }
  if((Pname==null)||(Pname=="")){
  if(("ProjectName".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  Pname = xlDriver.Value(colsList[idx+1]).toString().trim();
  }
  }
  
  if((Vname==null)||(Vname=="")){
  if(("VersionName".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  Vname = xlDriver.Value(colsList[idx+1]).toString().trim();    

  }
  }
  
  if((Cname==null)||(Cname=="")){
  if(("CycleName".indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  Cname = xlDriver.Value(colsList[idx+1]).toString().trim();    

  }
  }

  
//  testcase = "ALL";
//  OpcoNum = "1307";
  Lang = County();
//  Log.Message(Lang);
  
  
 
  }
xlDriver.Next();
 }
 
// path = "TestResource\\"+TestingType+"\\DS"+"_"+Lang+"_"+TestingType+".xlsx";
 if(TestingType.toUpperCase()=="SMOKE")
 path = "TestResource\\Smoke\\DS"+"_"+Lang+"_SMOKE.xlsx";
 else
 path = "TestResource\\Regression\\DS"+"_"+Lang+"_REGRESSION.xlsx";
 Log.Message(Project.Path+path)
 DDT.CloseDriver(xlDriver.Name);
 return path;    
  
}


function setPath(Region){ 
  Country = Region;
  Lang = County();
//  path = "TestResource\\"+TestingType+"\\DS"+"_"+Lang+"_"+TestingType+".xlsx";
  if(TestingType.toUpperCase()=="SMOKE")
   path = "TestResource\\Smoke\\DS"+"_"+Lang+"_SMOKE.xlsx";
  else
   path = "TestResource\\Regression\\DS"+"_"+Lang+"_REGRESSION.xlsx";
//  Log.Message(Project.Path+path)
}


function getBusinessFlow(){
// ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"EnvDetails",false);
// businessFlow = ExcelUtils.getRowData("Business_flow");;
//  return businessFlow;

//return "Default_Business_Flow";
return countrys();
}


function countrys(){ 
var temp = "";
var t_Type = "";

switch(EnvParams.Country.toUpperCase()) { // if we need to match case sensitive put Uppercase with in switch "baseCurrency.toUpperCase()"
case "INDIA":{
temp = "IND"
}
break;

case "CHINA":{
temp = "CHN"
}
break;

case "SINGAPORE":{
temp = "SNG"
}
break;
case "MALAYSIA":{
temp = "MLY"
}
break;

case "SPAIN":{
temp = "SPN"
}
break;

default:{
temp = ""; 
}
}



//if(EnvParams.TestingType.toLowerCase()=="full_regression")
//t_Type = "Regression";
//if(EnvParams.TestingType.toLowerCase()=="critical_regression")
//t_Type = "Critical Regression";
//if(EnvParams.TestingType.toLowerCase()=="sit")
//t_Type = "SIT";
//if(EnvParams.TestingType.toLowerCase()=="smoke")
//t_Type = "Smoke";
//
//
//if(EnvParams.TestingType.toLowerCase()=="smoke")
//temp = "Smoke";
//else if(EnvParams.TestingType.toLowerCase()=="critical_regression")
//temp = temp+"_"+t_Type;
//else if(EnvParams.TestingType.toLowerCase()=="sit")
//temp = temp+"_"+t_Type+"_FullCycle";
//else if(EnvParams.testcase!="ALL")
//temp = temp+"_"+t_Type+"_FullCycle";
//else
//temp = "GlobalTestPack";


if(EnvParams.TestingType.toLowerCase()=="full_regression")
t_Type = "Regression";
if(EnvParams.TestingType.toLowerCase()=="critical_regression")
t_Type = "Critical Regression";
if(EnvParams.TestingType.toLowerCase()=="sit")
t_Type = "Critical Regression";
if(EnvParams.TestingType.toLowerCase()=="smoke")
t_Type = "Smoke";


if(EnvParams.TestingType.toLowerCase()=="smoke")
temp = "Smoke";
else if(EnvParams.TestingType.toLowerCase()=="critical_regression")
temp = temp+"_"+t_Type;
else if(EnvParams.TestingType.toLowerCase()=="sit")
temp = temp+"_"+t_Type;
else if(EnvParams.TestingType.toLowerCase()=="full_regression")
temp = temp+"_"+t_Type;

return temp;

}




function LanChange(Languages){ 
var temp = "";
var t_Type = "";

switch(Languages.toUpperCase()) { // if we need to match case sensitive put Uppercase with in switch "baseCurrency.toUpperCase()"
case "ENGLISH":{
temp = "English"
}
break;

case "SPANISH":{
temp = "Spanish"
}
break;

case "CHINESE":{
temp = "Chinese (Simplified)"
}
break;

default:{
temp = ""; 
}
}
return temp;

}



function County(){ 
var temp = "";
var t_Type = "";
//Log.Message(EnvParams.Country.toUpperCase())
switch(EnvParams.Country.toUpperCase()) { // if we need to match case sensitive put Uppercase with in switch "baseCurrency.toUpperCase()"
case "INDIA":{
temp = "IND"
}
break;

case "CHINA":{
temp = "CHN"
}
break;

case "SINGAPORE":{
temp = "SNG"
}
break;
case "MALAYSIA":{
temp = "MLY"
}
break;

case "SPAIN":{
temp = "SPN"
}
break;

default:{
temp = ""; 
}
}
return temp;

}



function SetLanguage(Languages){ 
  Language = Languages;
}