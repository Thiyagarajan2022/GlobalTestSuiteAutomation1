

var excelName, sheet;
var excelObj;

function setExcelName(excelname, excelSheet){
excelName = excelname;
sheet = excelSheet;
excelObj = DDT.ExcelDriver(excelName,sheet,false);

}



//function aaa()
//{
//excelObj = DDT.ExcelDriver("C:\Users\Administrator\Desktop\TC_Craft\TC_Craft\TC_CraftFramework_V0.1\DataTables\TC_RunManager.xlsx","Business_Flow",false);
//var value;
//
//}

function getColumnValue( rowName, columnNum)
{
excelObj = DDT.ExcelDriver(excelName,sheet,false);
var value = '';
try{
excelObj.First();
for(var i=1; i!=0; i++){
    if(rowName==excelObj.ColumnName(0)){
        
      value = excelObj.ColumnName(columnNum);
      //excelObj.Value(columnNum);
      break;
        
    }
    else{
     excelObj.Next();
    }
     if(excelObj.EOF()){
        break;
     }
      
      
}
   
}catch( e){
Log.Message(e);
}


return value;

}

function getRowValue(columnName, rowNum)
{
var columnIndex;
var value ;

try{
excelObj.First();
//Log.Message(excelObj.ColumnCount)
 for(var j=0; j<excelObj.ColumnCount; j++){
       if(columnName==excelObj.ColumnName(j)){
       columnIndex = j;
//       Log.Message(columnIndex);
       break;
       }
 }

excelObj = DDT.ExcelDriver(excelName,sheet,false);
var i=1;
excelObj.First();
//Log.Message(excelObj.ColumnCount)
while(!excelObj.EOF()){

if(i==rowNum){
value = excelObj.Value(columnIndex);
//Log.Message(excelObj.Value(i))
//Log.Message("value :"+value)
}
//Log.Message(i)
i++;
excelObj.Next();
}

//   for(var i=0; i<=rowNum; i++){
//   value = excelObj.Value(columnIndex);
//  
//  excelObj.Next();
//  if(excelObj.EOF()){
//          break;
//       }
//      }

}catch(e){}


return value;

}


function getRowCount(){

var count =1;
excelObj.First();
while(!excelObj.EOF()){
excelObj.Next();
       count++;
     }
      
return count;
}

function getColumnCount(){
   
return excelObj.ColumnCount;
}


function getData( rowName, paramName)
{
excelObj = DDT.ExcelDriver(excelName,sheet,false);
var value = '';
var colNum = null;
try{
excelObj.First();

for(var j=0; j<excelObj.ColumnCount;j++){
 if(excelObj.ColumnName(j)==paramName)
 {
   colNum = j;
   break;
 }
 
}
for(var i=1; i!=0; i++){
 
 
    if(rowName==excelObj.Value(0)){
        
      value = excelObj.Value(colNum);
      //excelObj.Value(columnNum);
      break;
        
    }
    else{
     excelObj.Next();
    }
     if(excelObj.EOF()){
        break;
     }
      
      
}
   
}catch( e){
Log.Message(e);
}


return value;

}

function getRowData(rowidentifier)
{

var xlDriver = DDT.ExcelDriver(excelName,sheet,false);
var id =0;
var colsList = [];
 var temp ="";

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
//         Log.Message( colsList[idx]);
   }
     while (!DDT.CurrentDriver.EOF()) {
    
      for(var idx=0;idx<colsList.length-1;idx++){  
//      Log.Message(xlDriver.Value(colsList[idx])); 
      if(xlDriver.Value(colsList[idx]).toString().trim()==rowidentifier){
      temp = temp+xlDriver.Value(colsList[idx+1]).toString().trim();
//      Log.Message(temp);
      break;
      }
      }
    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return temp;
}  
 



