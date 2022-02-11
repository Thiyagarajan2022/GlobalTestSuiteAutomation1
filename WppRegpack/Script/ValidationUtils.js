﻿//USEUNIT ReportUtils
function verify(actual, expected, message){
if(actual == expected){
Log.Checkpoint(message);
ReportUtils.logStep("PASS",message);
}
else{  
TestRunner.JiraStat = false;
ReportUtils.logStep("FAIL",message);
Log.Error(message);
}
}
function getMenuItemsText(items)
{
var texts = [];
for(var i=items.length-1; i>=0; i--){
items[i].Caption;
texts.push( items[i].Caption);
}
return texts;
}

function compareList(actual, expected)
{
if(actual.length == expected.length){
for(var i=0; i<actual.length; i++){
if(actual[i]!=expected[i]){
  return false;
}
}
return true;
}
else{
return false;
}
}

