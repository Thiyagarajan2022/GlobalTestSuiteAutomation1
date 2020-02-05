﻿

function getPDFDocument(fileName)
{
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(Project.Path+"\\MPLReports"+"\\"+fileName);
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
  return docObj;
}

function getTextFromPDF(docObj){
 var textobj;
  try{
  obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj);
  Log.Message(textobj)
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  return textobj;
}