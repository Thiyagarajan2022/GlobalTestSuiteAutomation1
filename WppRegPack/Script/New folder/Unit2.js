function vv(){ 
  var docObj;

  // Load the PDF file to the PDDocument object
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3("C:\\Users\\674087\\Downloads\\p_jobinvoice-3.pdf");
  Job = SOXexcel(sheetName,1);
  var textobj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  
  Log.Message( "hi:"+textobj.getText_2(docObj)); 
//try{ 
  Log.Message("Job No: "+Job[0]);
  if(textobj.getText_2(docObj).contains("Job No: "+Job[0])){ 
      ValidationUtils.verify(true,true,"Job Number is available");
  } else{ 
   ValidationUtils.verify(true,false,"Job Number is Not-available in PDF"); 
  }
  
  Log.Message("Job Name: "+Job[1]);
  if(textobj.getText_2(docObj).contains("Job Name: "+Job[1])){ 
    ValidationUtils.verify(true,true,"Job Name is available");
  }
  else{ 
   ValidationUtils.verify(true,false,"Job Name is Not-available in PDF"); 
  }
  invoices = excel(sheetName,2);

  var count = 0;
  for(var i=0;i<invoices.length;i++){ 
    Log.Message("Invoice List: "+invoices[i]);
    if(textobj.getText_2(docObj).contains(invoices[i])){ 
      count++;
    }
  }
  if(count==invoices.length)
  ValidationUtils.verify(true,true,"Invoice details are available in PDF");
  else
  ValidationUtils.verify(true,false,"Invoice details are Not-available in PDF");
  
//  }
//catch(e){ 
//Log.Message("catch")
//  Log.Message(e);
//}
}