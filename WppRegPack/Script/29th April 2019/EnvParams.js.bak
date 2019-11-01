//USEUNIT ExcelUtils
//USEUNIT TextUtils
var instanceData = null;
var businessFlow = null;
function getEnvironment(){
var i;
var nArgs = BuiltIn.ParamCount();
var result = null;
var sheetLoc = null;
var colsList=[];
var instanceData;
var xlDriver= DDT.ExcelDriver(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"EnvDetails",false);

for (i = 1; i <= nArgs ; i++){    
if(BuiltIn.ParamStr(i).indexOf("runinstance")!=-1){
   var inst = BuiltIn.ParamStr(i);
   result = (inst.substring(inst.indexOf(":"))).trim();      
   break;
}
}
if(result==null){
  result =  "InstanceToRun";
}
for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
 colsList[idx] = DDT.CurrentDriver.ColumnName(idx);     
}
 while (!DDT.CurrentDriver.EOF()) {    
  for(var idx=0;idx<colsList.length-1;idx++){        
  if((result.indexOf(xlDriver.Value(colsList[idx]).toString().trim()))!=-1){
  instanceData = xlDriver.Value(colsList[idx+1]).toString().trim();    
  break;
  }
  }
xlDriver.Next();
 }
 DDT.CloseDriver(xlDriver.Name);
 return instanceData;    
  
}


function getBusinessFlow(){
 ExcelUtils.setExcelName(Project.Path+TextUtils.GetProjectValue("EnvDetailsPath"),"EnvDetails",false);
 businessFlow = ExcelUtils.getRowData("Business_flow");;
  return businessFlow;
}
