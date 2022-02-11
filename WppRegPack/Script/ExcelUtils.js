//USEUNIT EnvParams

var excelName, sheet,excelpath;
var excelObj;
var excelApp,wrkbook,columnCount,rowCount,excel_App,workbook;


function setExcelName(excelname, excelSheet){
excelName = excelname;
sheet = excelSheet;
JavaClasses.org_excelwrite.companyinfo.setExcelName(excelname, excelSheet);
//excelObj = DDT.ExcelDriver(excelName,sheet,true);
}

// Create Excel
function create_AutomationStat_Excel(fileName)
{
excelApp = Sys.OleObject("Excel.Application");
wrkbook = excelApp.Workbooks.Add();
excelApp.Visible = "false";
var text = EnvParams.Opco;
wrkbook.ActiveSheet.Cells.Item(1,1).Value="ModuleName";
wrkbook.ActiveSheet.Cells.Item(1,2).Value="TestCase_Name";
wrkbook.SaveAs(fileName);
wrkbook.Close();
excelApp.Quit();
}

// Publish values to workbook excel
function writeTo_AutomationStat_Excel(filePath,moduleName,testname,opco,executionTime)
{
  
filePath = TestRunner.automationStat_file
moduleName = TestRunner.moduleName
testname = TestRunner.JkinsName
opco = EnvParams.Opco
//executionTime = TestRunner.executionTime


  excel_App = Sys.OleObject("Excel.Application");
  workbook = excel_App.Workbooks.Open(filePath);
  sheet = workbook.Sheets.Item("Sheet1");
  columnCount = sheet.UsedRange.Columns.Count;
  rowCount = sheet.UsedRange.Rows.Count+1;
  
  var columnNo;
  var opcoHeadline = false;
  for(var x=columnCount; x>=columnCount-1;x--)
  {
    if(workbook.ActiveSheet.Cells.Item(1,x).Text.toString().trim().includes(opco))
    {
       columnNo = x;
       opcoHeadline = true;
       break;
    }      
  }     
  if(!opcoHeadline) 
  {   
      columnNo = columnCount+1;
      workbook.ActiveSheet.Cells.Item(1,columnNo).Value="Opco_"+opco+"_ExecutionTime(min)";
  }   

  var keyrow=0;
  var tcFound = false;
  for(var i=2;i<=rowCount;i++){
    if(workbook.ActiveSheet.Cells.Item(i,2).Text.toString().trim() == testname.trim())
    {
      if(workbook.ActiveSheet.Cells.Item(i,columnNo).Text=="")
      {
          workbook.ActiveSheet.Cells.Item(i,columnNo).Value=executionTime;
          tcFound = true;
          break;
      }
     }
  }
  if(!tcFound)
  {
      workbook.ActiveSheet.Cells.Item(rowCount,1).Value=moduleName;
      workbook.ActiveSheet.Cells.Item(rowCount,2).Value=testname;
      workbook.ActiveSheet.Cells.Item(rowCount,columnNo).Value=executionTime;
  
  }
//  if(tcFound)
//    {
//      if(workbook.ActiveSheet.Cells.Item(keyrow,columnNo).Text=="")
//        workbook.ActiveSheet.Cells.Item(keyrow,columnNo).Value=executionTime;
//      else
//         workbook.ActiveSheet.Cells.Item(rowCount,columnNo).Value=executionTime;
//    }
//    else
//    {
//     workbook.ActiveSheet.Cells.Item(rowCount,1).Value=moduleName;
//      workbook.ActiveSheet.Cells.Item(rowCount,2).Value=testname;
//      workbook.ActiveSheet.Cells.Item(rowCount,columnNo).Value=executionTime;
//    }
  workbook.Save();
  //workbook.Close();
}

// Close Excel
function close_AutomationStat_Excel()
{
  excel_App.Quit();
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
 for(var j=0; j<excelObj.ColumnCount; j++){
       if(columnName==excelObj.ColumnName(j)){
       columnIndex = j;
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

var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";
if(rowidentifier.indexOf("Central Team - Client Management")!=-1)
 rowidentifier = "Central Team - Client Account Management";
if(rowidentifier.indexOf("Central Team - Vendor Management")!=-1)
 rowidentifier = "Central Team - Vendor Account Management";
if(rowidentifier.indexOf("SSC - Expense Cashiers")!=-1)
 rowidentifier = "SSC - Cashier";
if(rowidentifier.indexOf("SSC - Billers")!=-1)
 rowidentifier = "SSC - Biller";
 
     while (!DDT.CurrentDriver.EOF()) {

       if(xlDriver.Value(0).toString().trim()==rowidentifier){
        try{
         temp = temp+xlDriver.Value(column).toString().trim();
         }
        catch(e){
        temp = "";
        }

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

}
     while (!DDT.CurrentDriver.EOF()) {

       if(xlDriver.Value(Col).toString().trim().indexOf(rowidentifier.toString().trim())!=-1){
        try{
         temp = temp+xlDriver.Value(Col).toString().trim();
         }
        catch(e){
        temp = "";
        }
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
var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
var id =0;
var colsList = [];
 var temp ="";
if(rowidentifier.indexOf("OpCo -")!=-1){ 
  rowidentifier = rowidentifier.replace(/OpCo -/g,column);
  }
if(rowidentifier.indexOf("Billers")!=-1)
    rowidentifier = rowidentifier.replace(/Billers/g,"Biller");
if((rowidentifier.indexOf("(")!=-1)&&(rowidentifier.indexOf(")")!=-1))
    rowidentifier = rowidentifier.substring(0,rowidentifier.indexOf("(")-1);
    
var Col = "";
for(var i=0;i<DDT.CurrentDriver.ColumnCount;i++){ 
  if(DDT.CurrentDriver.ColumnName(i).toString().trim().indexOf(column)!=-1)
  Col = DDT.CurrentDriver.ColumnName(i).toString().trim();
}
     while (!DDT.CurrentDriver.EOF()) {
       if(xlDriver.Value(Col).toString().trim().indexOf(rowidentifier.toString().trim())!=-1){
        try{
         temp = temp+xlDriver.Value(Col).toString().trim();
         }
        catch(e){
        temp = "";
        }
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
var temp = JavaClasses.org_excelwrite.companyinfo.getColumnDatas(rowidentifier,column)
//Log.Message("temp :"+temp);
if((temp!="")&&(temp!=null)){
return temp.OleValue.toString().trim();
}else{ 
  return "";
}

////Log.Message("excelName :"+excelName)
////Log.Message("sheet :"+sheet);
//var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
//var id =0;
//var colsList = [];
// var temp ="";
////Log.Message("rowidentifier :"+rowidentifier)
////Log.Message("column :"+column);
//     while (!DDT.CurrentDriver.EOF()) {
////    Log.Message("Colunm :"+xlDriver.Value(0).toString().trim())
//       if(xlDriver.Value(0).toString().trim()==column){
//        try{
//         temp = temp+xlDriver.Value(rowidentifier).toString().trim();
//         }
//        catch(e){
//        temp = "";
//        }
////      Log.Message("temp :"+temp);
//      break;
//      }
//
//    xlDriver.Next();
//     }
//     DDT.CloseDriver(xlDriver.Name);
//Log.Message("temp :"+temp);
//     return temp;
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
var temp = "";
setExcelName(excelName, sheets, true);
temp = JavaClasses.org_excelwrite.companyinfo.getRowDatas(array,Opco);


/*

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
 
 */
 
 
if((temp!="")&&(temp!=null)){
return temp.OleValue.toString().trim();
}else{ 
  return "";
}
 
}



function WriteExcelSheet(array,Opco,sheets,val){
  
JavaClasses.org_excelwrite.companyinfo.WriteExcelSheet(array,Opco,sheets,val)

////Log.Message("Execution completed,sending result to excel book , FileName:"+excelName+"  sheetname:"+sheets);
////  Log.Message(val);
////  Log.Message(array)
//  var app = Sys.OleObject("Excel.Application");
//  app.Visible = "True";
//  var curArrayVals = [];  
//  var book = app.Workbooks.Open(excelName);
////  Log.Message(sheets)
//  var sheet = book.Sheets.Item(sheets);;
//  var columnCount = sheet.UsedRange.Columns.Count;
//  var rowCount = sheet.UsedRange.Rows.Count;
////  Log.Message(columnCount);
////  Log.Message(rowCount);
//  var arrays={};
//  var idx =0;
//  var col =0;
//  var row = 0;
//  for(var k = 1; k<=columnCount;k++){
//  if(sheet.Cells.Item(1, k).Text.toString().trim()==Opco){
//  col = k;
////  Log.Message(sheet.Cells.Item(1, k).Text)
//  }
//  }
//  var rowStatus = false;
//  for(var k = 1; k<=rowCount;k++){
//  if(sheet.Cells.Item(k, 1).Text.toString().trim()==array){
////  Log.Message(sheet.Cells.Item(k, 1).Text);
////  Log.Message(sheet.Cells.Item(k, col).Text);
//  row = k;
//  rowStatus = true;
//  }
//  }
//  
//  if(!rowStatus){ 
//   sheet.Cells.Item(rowCount+1,  1).Value = array
//   sheet.Cells.Item(rowCount+1,  col).Value = val
////   Log.Message("Row :"+rowCount)
////   Log.Message("Column :"+col)
//  }
//  else{ 
//    sheet.Cells.Item(row,  col).Value = val
////  Log.Message("Row :"+row)
////   Log.Message("Column :"+col)
//  }
// book.Save();
// app.Quit();


}




