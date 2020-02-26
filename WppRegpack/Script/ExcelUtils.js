

var excelName, sheet;
var excelObj;

function setExcelName(excelname, excelSheet){
excelName = excelname;
sheet = excelSheet;
JavaClasses.org_excelwrite.companyinfo.setExcelName(excelname, excelSheet);
//Log.Message("excelName :"+excelName);
//Log.Message("sheet :"+sheet);
//excelObj = DDT.ExcelDriver(excelName,sheet,true);

}



function getColumnValue( rowName, columnNum)
{
excelObj = DDT.ExcelDriver(excelName,sheet,true);
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

DDT.CloseDriver(excelObj.Name)
return value;

}

function getRowValue(columnName, rowNum)
{
var columnIndex;
var value ;
excelObj = DDT.ExcelDriver(excelName,sheet,true);
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

excelObj = DDT.ExcelDriver(excelName,sheet,true);
var i=1;
excelObj.First();
while(!excelObj.EOF()){

if(i==rowNum){
value = excelObj.Value(columnIndex);
}
i++;
excelObj.Next();
}

}catch(e){}

DDT.CloseDriver(excelObj.Name)
return value;

}


function getRowCount(){
excelObj = DDT.ExcelDriver(excelName,sheet,true);
var count =1;
excelObj.First();
while(!excelObj.EOF()){
excelObj.Next();
       count++;
     }
DDT.CloseDriver(excelObj.Name)
return count;
}

function getColumnCount(){
   
return excelObj.ColumnCount;
}


function getData( rowName, paramName)
{
excelObj = DDT.ExcelDriver(excelName,sheet,true);
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

DDT.CloseDriver(excelObj.Name)
return value;

}

function getRowData(rowidentifier)
{

var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
//         Log.Message( colsList[idx]);
   }
     while (!DDT.CurrentDriver.EOF()) {
    
      for(var idx=0;idx<colsList.length-1;idx++){  
       Log.Message(xlDriver.Value(colsList[idx]))
       Log.Message(rowidentifier)
      if(xlDriver.Value(colsList[idx]).toString().trim()==rowidentifier){
      temp = temp+xlDriver.Value(colsList[idx+1]).toString().trim();

      break;
      }
      }
    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return temp;
}  
 


function getRowData1(rowidentifier)
{

var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
         
   }
     while (!DDT.CurrentDriver.EOF()) {
    
      for(var idx=0;idx<colsList.length-1;idx++){  
        if(xlDriver.Value(colsList[idx]).toString().trim()==rowidentifier){
        try{
         temp = temp+xlDriver.Value(colsList[idx+1]).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }
      }
    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return temp;
}
 

function getRowDatas(rowidentifier,column)
{
var temp = JavaClasses.org_excelwrite.companyinfo.getRowDatas(rowidentifier,column);

////Log.Message("excelName :"+excelName);
////Log.Message("sheet :"+sheet);
////Log.Message("column :"+column);
//var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
//var id =0;
//var colsList = [];
// var temp ="";
//
//     while (!DDT.CurrentDriver.EOF()) {
////    Log.Message("Colunm :"+xlDriver.Value(0).toString().trim())
//       if(xlDriver.Value(0).toString().trim()==rowidentifier){
//        try{
//          
//         temp = temp+xlDriver.Value(column).toString().trim();
//         }
//        catch(e){
//        temp = "";
//        }
//
////      Log.Message("temp :"+temp);
//      break;
//      }
////      Log.Message("temp :"+temp);
//    xlDriver.Next();
//     }
//     DDT.CloseDriver(xlDriver.Name);
     
if((temp!="")&&(temp!=null)){
return temp.OleValue.toString().trim();
}else{ 
  return "";
}
}



function SSCLogin(rowidentifier,column)
{
//Log.Message("excelName :"+excelName);
//Log.Message("sheet :"+sheet);
var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message(rowidentifier)
//Log.Message(column)

if(rowidentifier.indexOf("Central Team - Client Management")!=-1)
 rowidentifier = "Central Team - Client Account Management";
if(rowidentifier.indexOf("Central Team - Vendor Management")!=-1)
 rowidentifier = "Central Team - Vendor Account Management";
if(rowidentifier.indexOf("SSC - Expense Cashiers")!=-1)
 rowidentifier = "SSC - Cashier";

     while (!DDT.CurrentDriver.EOF()) {
//    Log.Message("Colunm :"+xlDriver.Value(2).toString().trim())
       if(xlDriver.Value(0).toString().trim()==rowidentifier){
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message("temp :"+temp);
      break;
      }

    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     
     if((temp=="")||(temp==null)){ 
if((rowidentifier.indexOf("(")!=-1)&&(rowidentifier.indexOf(")")!=-1))
    rowidentifier = rowidentifier.substring(0,rowidentifier.indexOf("(")-1);
id =0;
colsList = [];
  xlDriver = DDT.ExcelDriver(excelName,sheet,true);
  var Col = "";
  for(var i=0;i<DDT.CurrentDriver.ColumnCount;i++){ 
  if(DDT.CurrentDriver.ColumnName(i).toString().trim().indexOf(column)!=-1)
  Col = DDT.CurrentDriver.ColumnName(i).toString().trim();
//  else
//  Log.Message(DDT.CurrentDriver.ColumnName(i).toString().trim())
}
     while (!DDT.CurrentDriver.EOF()) {
//    Log.Message("Colunm :"+xlDriver.Value(Col).toString().trim())
       if(xlDriver.Value(Col).toString().trim().indexOf(rowidentifier.toString().trim())!=-1){
        try{
         temp = temp+xlDriver.Value(Col).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message("temp :"+temp);
      break;
      }

    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     }
     
     
     
     
     
     
     
     
     return temp;
}



function AgencyLogin(rowidentifier,column)
{
//Log.Message("excelName :"+excelName);
//Log.Message("sheet :"+sheet);
var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";
if(rowidentifier.indexOf("OpCo -")!=-1){ 
  rowidentifier = rowidentifier.replace(/OpCo -/g,column);
  }
if(rowidentifier.indexOf("Billers")!=-1)
    rowidentifier = rowidentifier.replace(/Billers/g,"Biller");
//Log.Message(rowidentifier)    
if((rowidentifier.indexOf("(")!=-1)&&(rowidentifier.indexOf(")")!=-1))
    rowidentifier = rowidentifier.substring(0,rowidentifier.indexOf("(")-1);
    
var Col = "";
for(var i=0;i<DDT.CurrentDriver.ColumnCount;i++){ 
  if(DDT.CurrentDriver.ColumnName(i).toString().trim().indexOf(column)!=-1)
  Col = DDT.CurrentDriver.ColumnName(i).toString().trim();
//  else
//  Log.Message(DDT.CurrentDriver.ColumnName(i).toString().trim())
}
//Log.Message(rowidentifier)
//Log.Message(Col)
//Log.Message(column)
     while (!DDT.CurrentDriver.EOF()) {
//    Log.Message("Colunm :"+xlDriver.Value(Col).toString().trim())
       if(xlDriver.Value(Col).toString().trim().indexOf(rowidentifier.toString().trim())!=-1){
        try{
         temp = temp+xlDriver.Value(Col).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message("temp :"+temp);
      break;
      }

    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return temp;
}




function getAllRowDatas(column,row)
{
var temp = JavaClasses.org_excelwrite.companyinfo.getAllRowDatas(column,row)
////Log.Message("excelName :"+excelName);
////Log.Message("sheet :"+sheet);
//var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
//var id =0;
//var colsList = [];
// var temp ="";
////Log.Message(rowidentifier)
////Log.Message(column)
////Log.Message(row)
//var i=0;
//     while (!DDT.CurrentDriver.EOF()) {
////    Log.Message("Colunm :"+xlDriver.Value(0).toString().trim())
////Log.Message(row)
//       if(i==row)
//       if((xlDriver.Value(column).toString().trim()!="")&&(xlDriver.Value(column).toString().trim()!=null)){
//        try{
//         temp = temp+xlDriver.Value(column).toString().trim();
//         }
//        catch(e){
//        temp = "";
//        }
////      Log.Message("temp :"+temp);
//      break;
//      }
//
//    xlDriver.Next();
//    i++;
//     }
//     DDT.CloseDriver(xlDriver.Name);
//     return temp;
//Log.Message("Temp :"+temp)
if((temp!="")&&(temp!=null)){
return temp.OleValue.toString().trim();
}else{ 
  return "";
}
}





function getColumnDatas(rowidentifier,column)
{
//Log.Message("excelName :"+excelName)
//Log.Message("sheet :"+sheet);
var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";
//Log.Message("rowidentifier :"+rowidentifier)
//Log.Message("column :"+column);
     while (!DDT.CurrentDriver.EOF()) {
//    Log.Message("Colunm :"+xlDriver.Value(0).toString().trim())
       if(xlDriver.Value(0).toString().trim()==column){
        try{
         temp = temp+xlDriver.Value(rowidentifier).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message("temp :"+temp);
      break;
      }

    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return temp;
}


function getRowDatas_(rowidentifier,column)
{

var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";

     while (!DDT.CurrentDriver.EOF()) {
    
       if(xlDriver.Value(colsList[column]).toString().trim()==rowidentifier){
        try{
         temp = temp+xlDriver.Value(colsList[column+1]).toString().trim();
         }
        catch(e){
        temp = "";
        }
//      Log.Message(temp);
      break;
      }

    xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return temp;
}




function ReadExcelSheet(array,Opco,sheets){
var temp = ""

//Log.Message("Execution completed,sending result to excel book , FileName:"+excelName+"sheetname:"+sheet);
  var app = Sys.OleObject("Excel.Application");
//  app.Visible = "True";
  var curArrayVals = [];  
  var book = app.Workbooks.Open(excelName);
  var sheet = book.Sheets.Item(sheets);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;
//  Log.Message(columnCount);
//  Log.Message(rowCount);
  var arrays={};
  var idx =0;
  var col =0;
  var row = 0;
  for(var k = 1; k<=columnCount;k++){
  if(sheet.Cells.Item(1, k).Text.toString().trim()==Opco){
  col = k;
  }
  }
  var rowStatus = false;
  for(var k = 1; k<=rowCount;k++){
  if(sheet.Cells.Item(k, 1).Text.toString().trim()==array){
  row = k;
  rowStatus = true;
  }
  }
  
  if(rowStatus){ 
   temp = sheet.Cells.Item(row,  col).Text;
//   Log.Message(temp)
  }
// book.Save();
 app.Quit();
 return temp;
}



function WriteExcelSheet(array,Opco,sheets,val){
//Log.Message("Execution completed,sending result to excel book , FileName:"+excelName+"  sheetname:"+sheets);
//  Log.Message(val);
//  Log.Message(array)
  var app = Sys.OleObject("Excel.Application");
  app.Visible = "True";
  var curArrayVals = [];  
  var book = app.Workbooks.Open(excelName);
//  Log.Message(sheets)
  var sheet = book.Sheets.Item(sheets);;
  var columnCount = sheet.UsedRange.Columns.Count;
  var rowCount = sheet.UsedRange.Rows.Count;
//  Log.Message(columnCount);
//  Log.Message(rowCount);
  var arrays={};
  var idx =0;
  var col =0;
  var row = 0;
  for(var k = 1; k<=columnCount;k++){
  if(sheet.Cells.Item(1, k).Text.toString().trim()==Opco){
  col = k;
//  Log.Message(sheet.Cells.Item(1, k).Text)
  }
  }
  var rowStatus = false;
  for(var k = 1; k<=rowCount;k++){
  if(sheet.Cells.Item(k, 1).Text.toString().trim()==array){
//  Log.Message(sheet.Cells.Item(k, 1).Text);
//  Log.Message(sheet.Cells.Item(k, col).Text);
  row = k;
  rowStatus = true;
  }
  }
  
  if(!rowStatus){ 
   sheet.Cells.Item(rowCount+1,  1).Value = array
   sheet.Cells.Item(rowCount+1,  col).Value = val
//   Log.Message("Row :"+rowCount)
//   Log.Message("Column :"+col)
  }
  else{ 
    sheet.Cells.Item(row,  col).Value = val
//  Log.Message("Row :"+row)
//   Log.Message("Column :"+col)
  }
 book.Save();
 app.Quit();
}


//Strating Of TestCase

