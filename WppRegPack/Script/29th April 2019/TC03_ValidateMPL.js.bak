//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT PdfUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "Invoice";
var values = [];
var invoices = [];

function validateMPL(){
   var docObj;
   values = SOXexcel(sheetName,1);
   docObj = PdfUtils.getPDFDocument(values[1]);
   var txtbj = PdfUtils.getTextFromPDF(docObj);
   ValidationUtils.verify(txtbj.contains("Job No: "+values[2]),true,"Job Number is available");
   ValidationUtils.verify(txtbj.contains("Job Name: "+values[3]),true,"Job Name is available");
   invoices = excel(sheetName,2);

  var count = 0;
  for(var i=0;i<invoices.length;i++){    
    if(txtbj.contains(invoices[i])){ 
      count++;
    }
  }
  if(count==invoices.length)
  ValidationUtils.verify(true,true,"Invoice details are available in PDF");
  else
  ValidationUtils.verify(true,false,"Invoice details are Not-available in PDF");

}


function SOXexcel(CreateClient,start){ 
var Arrayss = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
     Arrayss[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrayss;
}


function excel(CreateClient,start){ 

var Arrayss = [];
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];
//Log.Message(DDT.CurrentDriver.ColumnCount);
   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();

     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=start;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim();
      if(idx!=colsList.length-1)
      temp = temp+" ";
      }
      else{ 
        temp = temp+" ";
      }
      }
     if(temp.length!=6){
     Arrayss[id]=temp;
//     Log.Message(temp)
     }
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return Arrayss;
}
