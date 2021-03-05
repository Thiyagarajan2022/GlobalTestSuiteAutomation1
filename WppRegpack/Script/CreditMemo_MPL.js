//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT WorkspaceUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var Language = "";

function CreditMemo()
{
//  var fileName = "C:\\Users\\674087\\Music\\July Regression\\GlobalTestSuiteAutomation\\WppRegPack\\MPLReports\\Critical_Regression\\India\\1707\\PDF\\Print Job Quote-23.pdf";
  var fileName = "";
    Language = "";
  Language = EnvParams.LanChange(EnvParams.Language);
  WorkspaceUtils.Language = Language;  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("PDF Credit Note",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"Credit Note PDF is needed to validate");
  }
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName);
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
//  var workBook = "C:\\Users\\674087\\Music\\July Regression\\GlobalTestSuiteAutomation\\WppRegpack\\TestResource\\Regression\\DS_IND_REGRESSION.xlsx";
  var sheetName = "CreditMemo";
//  EnvParams.Country = "India";
//  EnvParams.Opco = "1707";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
 
  var pdflineSplit = docObj.split("\r\n");
  
 
  var street = ReadExcelSheet("Street 1",EnvParams.Opco,"CreateClient");
  var postCode = ReadExcelSheet("Post Code",EnvParams.Opco,"CreateClient");
  var postDistrict = ReadExcelSheet("Post District",EnvParams.Opco,"CreateClient");
  var country = ReadExcelSheet("Country",EnvParams.Opco,"CreateClient");
  var Attn = ReadExcelSheet("Attn.",EnvParams.Opco,"CreateClient");
  var TaxNo = ReadExcelSheet("Tax No.",EnvParams.Opco,"CreateClient");
  ExcelUtils.setExcelName(workBook, "CreditMemoMPL", true);
//  var Revesion = ExcelUtils.getColumnDatas("Quote Revision",EnvParams.Opco)
  var GrandTotal = ExcelUtils.getColumnDatas("Invoice TOTAL",EnvParams.Opco)
  var PaymentTerm = ExcelUtils.getColumnDatas("Payment Terms",EnvParams.Opco)
  
ExcelUtils.setExcelName(workBook, "Data Management", true);
var clientName = ExcelUtils.getRowDatas("Global Client Name",EnvParams.Opco)
if((clientName=="")||(clientName==null)){
clientName = ReadExcelSheet("Client Name",EnvParams.Opco,"CreateClient");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
var productName = ExcelUtils.getRowDatas("Global Product Name",EnvParams.Opco)
if((productName=="")||(productName==null)){
productName = ReadExcelSheet("Product Name",EnvParams.Opco,"CreateClient");
}

   if((EnvParams.Country.toUpperCase()=="INDIA") || (EnvParams.Country.toUpperCase()=="SINGAPORE"))
   var index = pdflineSplit.indexOf("TAX INVOICE");
   else if((EnvParams.Country.toUpperCase()=="SPAIN") || (EnvParams.Country.toUpperCase()=="MALAYSIA")|| (EnvParams.Country.toUpperCase()=="CHINA"))
   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "CREDIT NOTE").OleValue.toString().trim());
//   else if(EnvParams.Country.toUpperCase()=="SINGAPORE")
//   var index = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "QUOTATION").OleValue.toString().trim());
   
       if(index>=0){
          ReportUtils.logStep("INFO","Heading is available Pdf")
          ValidationUtils.verify(true,true,"Heading is available Pdf")
          TextUtils.writeLog("Heading is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Heading is available Pdf")
          
    if(index>=0){
          ReportUtils.logStep("INFO",clientName+"ClientName is matching with Pdf")
          ValidationUtils.verify(true,true,clientName+" ClientName is matching with Pdf")
          TextUtils.writeLog(clientName+" ClientName is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"ClientName is not same in Invoice");
          if(EnvParams.Country.toUpperCase()=="INDIA"){
   var index = pdflineSplit.indexOf(productName);
    if(index>=0){
          ReportUtils.logStep("INFO",productName+" ProductName is matching with Pdf")
          ValidationUtils.verify(true,true,productName+" ProductName is matching with Pdf")
          TextUtils.writeLog(productName+" ProductName is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"ProductName is not same in Invoice"); 
          }
                
   var index = pdflineSplit.indexOf(street);
    if(index>=0){
          ReportUtils.logStep("INFO",street+" Street is matching with Pdf")
          ValidationUtils.verify(true,true,street+" Street is matching with Pdf")
          TextUtils.writeLog(street+" Street is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Street is not same in Credit Memo");
   var index = pdflineSplit.indexOf(postCode+"  "+postDistrict);
    if (index == -1)
        index = pdflineSplit.indexOf(postCode+" "+postDistrict);
    if(index>=0){
          ReportUtils.logStep("INFO",postCode+" "+postDistrict+" PostCode and Post District is matching with Pdf")
          ValidationUtils.verify(true,true,postCode+" "+postDistrict+" PostCode and Post District are matching with Pdf")
          TextUtils.writeLog(postCode+" "+postDistrict+" PostCode and Post District are matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"PostCode and Post District are not same in Invoice");
   var index = pdflineSplit.indexOf(country);
    if(index>=0){
          ReportUtils.logStep("INFO",country+" Country is matching with Pdf")
          ValidationUtils.verify(true,true,country+" Country is matching with Pdf")
          TextUtils.writeLog(country+" Country is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Country is not same in Draft Invoice");
   
  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
  if((jobNumber=="")||(jobNumber==null)){
  ExcelUtils.setExcelName(workBook, sheetName, true);
  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
  }
  if((jobNumber=="")||(jobNumber==null))
  ValidationUtils.verify(false,true,"Job Number is needed to Validate Credit Memo");
   
  var j, x, pdfJobNum, pointer, pdfJobName;
  
  var jobName = ReadExcelSheet("Job_name",EnvParams.Opco,"JobCreation");
ExcelUtils.setExcelName(workBook, "Data Management", true);
var JobCurrency = ExcelUtils.getRowDatas("Client Currency",EnvParams.Opco)
if((JobCurrency=="")||(JobCurrency==null)){
JobCurrency = ReadExcelSheet("Currency",EnvParams.Opco,"CreateClient");
}
  var productNumber = ReadExcelSheet("Global Product Number",EnvParams.Opco,"Data Management");
  var clientNumber = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
var pName = false;  
 Log.Message(pdflineSplit.length)
  for (j=0; j<pdflineSplit.length; j++)
  {
//    Log.Message(pdflineSplit[j])
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Attn.").OleValue.toString().trim()))
    {
      x= pdflineSplit[j].split(".");
      pdfJobNum = x[1].trim();
       if(Attn!=pdfJobNum)
        ValidationUtils.verify(false,true,"Attention is not same in Draft Invoice");
        else{
          ReportUtils.logStep("INFO",Attn+" Attention is matching with Pdf");
        ValidationUtils.verify(true,true,Attn+" Attention is matching with Pdf");
        TextUtils.writeLog(Attn+" Attention is matching with Pdf");
        }
    }

        if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Note No").OleValue.toString().trim()))
    {
      
    if((EnvParams.Country.toUpperCase()=="CHINA")&&(Language=="Chinese (Simplified)")){
      var atSize = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Credit Note No").OleValue.toString().trim();
      Log.Message("atSize :"+atSize)
      pdflineSplit[j] = pdflineSplit[j].substring(atSize.length+1); 
      x= pdflineSplit[j].split(" ");
      x[0]= pdflineSplit[j];
      x[1]= pdflineSplit[j];
      Log.Message("x[1] :"+x[1])
      }else
      x= pdflineSplit[j].split(":");
      pdfJobNum = x[1].trim();
       if(pdfJobNum.indexOf(EnvParams.Opco)==-1)
        ValidationUtils.verify(false,true,"Credit Note No is not same in Credit Memo");
        else{
        ReportUtils.logStep("INFO","Credit Note No is availble in Pdf")
        ValidationUtils.verify(true,true,"Credit Note No is availble in Pdf")
        TextUtils.writeLog("Credit Note No is availble in Pdf")
        }
    }
/*    
        if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Credit Note Date").OleValue.toString().trim()))
    {
      x= pdflineSplit[j].split(":");
      pdfJobNum = x[1].trim();
      var dateType = /(\d{1,2})([\/-])(\d{1,2})\2(\d{4})/;
      var isMatch = dateType.test(pdfJobNum);
      if(isMatch)
//       if("xxxxxx"==pdfJobNum)
        ValidationUtils.verify(false,true,"Invoice Date is not available in Draft Invoice");
        else{
        ReportUtils.logStep("INFO","Invoice Date is available in Pdf")
        ValidationUtils.verify(true,true,"Invoice Date is available in Pdf")
        TextUtils.writeLog("Invoice Date is available in Pdf")
        }
    }
    
        if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Due Date").OleValue.toString().trim()))
    {
      x= pdflineSplit[j].split(":");
      pdfJobNum = x[1].trim();
      var dateType = /(\d{1,2})([\/-])(\d{1,2})\2(\d{4})/;
      var isMatch = dateType.test(pdfJobNum);
      if(isMatch)
//       if("xxxxxx"!=pdfJobNum)
        ValidationUtils.verify(false,true,"Due Date is not available in Draft Invoice");
        else{
        ReportUtils.logStep("INFO","Due Date is available in Pdf")
        ValidationUtils.verify(true,true,"Due Date is available in Pdf")
        TextUtils.writeLog("Due Date is available in Pdf")
        }
    }
 */   
        if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Terms").OleValue.toString().trim()))
    {
      
    
    if((EnvParams.Country.toUpperCase()=="CHINA")&&(Language=="Chinese (Simplified)")){
        var atSize = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Payment Terms").OleValue.toString().trim();
      pdflineSplit[j] = pdflineSplit[j].substring(atSize.length+1); 
      x= pdflineSplit[j].split(" ");
      x[0]= pdflineSplit[j];
      x[1]= pdflineSplit[j];
      }else
      x= pdflineSplit[j].split(":");
      pdfJobNum = x[1].trim();
      Log.Message(PaymentTerm)
      Log.Message(pdfJobNum)
       if(pdfJobNum.indexOf(PaymentTerm)==-1)
        ValidationUtils.verify(false,true,"Payment Terms is not same in Draft Invoice");
        else{
        ReportUtils.logStep("INFO",PaymentTerm+" Payment Terms is matching with Pdf")
        ValidationUtils.verify(true,true,PaymentTerm+" Payment Terms is matching with Pdf")
        TextUtils.writeLog(PaymentTerm+" Payment Terms is matching with Pdf")
        }
    }
    
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job No").OleValue.toString().trim()))
    {
      
    
    if((EnvParams.Country.toUpperCase()=="CHINA")&&(Language=="Chinese (Simplified)")){
        var atSize = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Job No").OleValue.toString().trim();
      pdflineSplit[j] = pdflineSplit[j].substring(atSize.length+1); 
      x= pdflineSplit[j].split(" ");
      x[0]= pdflineSplit[j];
      x[1]= pdflineSplit[j];
      }else
      x= pdflineSplit[j].split(":");
      pdfJobNum = x[1].trim();
       if(jobNumber!=pdfJobNum)
        ValidationUtils.verify(false,true,"Job Number is not same in Invoice");
        else{
        ReportUtils.logStep("INFO",jobNumber+" Job Number is matching with Pdf")
        ValidationUtils.verify(true,true,jobNumber+" Job Number is matching with Pdf")
        TextUtils.writeLog(jobNumber+" Job Number is matching with Pdf")
        }
    }
//     if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Name").OleValue.toString().trim()))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobName = x[1].trim();
//      if(pdfJobName.includes(jobName)){
//          ReportUtils.logStep("INFO",jobName+" Job Name is matching with Pdf")
//          ValidationUtils.verify(true,true,jobName+" Job Name is matching with Pdf")
//          TextUtils.writeLog(jobName+" Job Name is matching with Pdf")
//          }
//          else
//          ValidationUtils.verify(false,true,"Job Name is not same in Invoice");
//    }
//    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job Currency").OleValue.toString().trim()))
//    {
//        if(pdflineSplit[j].includes(JobCurrency))
//         {
//          ReportUtils.logStep("INFO","Job Currency is matching with Pdf")
//          ValidationUtils.verify(true,true,JobCurrency+" Job Currency is matching with Pdf")
//          TextUtils.writeLog(JobCurrency+" Job Currency is matching with Pdf")
//          }
//          else{
//          ValidationUtils.verify(false,true,"Job Currency is not same in Invoice");
//        }
//    }
//
//     if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Product Number").OleValue.toString().trim()))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobName = x[1].trim();
//        if(pdflineSplit[j].includes(productNumber))
//         {
//          ReportUtils.logStep("INFO",productNumber+" Product Number is matching with Pdf")
//          ValidationUtils.verify(true,true,productNumber+" Product Number is matching with Pdf")
//          TextUtils.writeLog(productNumber+" Product Number is matching with Pdf")
//          }
//          else{
//          ValidationUtils.verify(false,true,"Product Number is not same in Invoice");
//        }
//    }
//
//     if(pdflineSplit[j].includes("Product Name"))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobName = x[1].trim();
//      Log.Message(pdflineSplit[j])
//      Log.Message(productName)
//        if(pdflineSplit[j].includes(productName))
//         {
//          ReportUtils.logStep("INFO",productName+" Product Name is matching with Pdf")
//          ValidationUtils.verify(true,true,productName+" Product Name is matching with Pdf")
//          TextUtils.writeLog(productName+" Product Name is matching with Pdf")
//          pName = true;
//          }
//          else{
//          ValidationUtils.verify(false,true,"Product Name is not same in Invoice");
//        }
//    }



var TaxVariable = "";
   if((EnvParams.Country.toUpperCase()=="SPAIN") || (EnvParams.Country.toUpperCase()=="MALAYSIA") || (EnvParams.Country.toUpperCase()=="CHINA")){
     Log.Message(pdflineSplit)
     Log.Message(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Tax No").OleValue.toString().trim())
   TaxVariable = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Tax No").OleValue.toString().trim());
        Log.Message(TaxVariable)
   }
   else if(EnvParams.Country.toUpperCase()=="SINGAPORE"){
   TaxVariable = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GST No").OleValue.toString().trim());
   }
//   Log.Message(TaxVariable)
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, TaxVariable).OleValue.toString().trim()))
    {
      
    
    if((EnvParams.Country.toUpperCase()=="CHINA")&&(Language=="Chinese (Simplified)")){
        var atSize = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Tax No").OleValue.toString().trim();
      pdflineSplit[j] = pdflineSplit[j].substring(atSize.length+1); 
      x= pdflineSplit[j].split(" ");
      x[0]= pdflineSplit[j];
      x[1]= pdflineSplit[j];
      }else
      x= pdflineSplit[j].split(":");
      pdfJobName = x[1].trim();
      Log.Message(pdflineSplit[j])
      Log.Message(TaxNo)
      Log.Message(pdflineSplit[j].includes(TaxNo))
        if(pdflineSplit[j].includes(TaxNo))
         {
          ReportUtils.logStep("INFO","Tax No is matching with Pdf")
          ValidationUtils.verify(true,true,TaxNo+" Tax No is matching with Pdf")
          TextUtils.writeLog(TaxNo+" Tax No is matching with Pdf")
          }
          else{
          ValidationUtils.verify(false,true,"Tax No is not same in Credit Memo");
        }
    }


     if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Client No").OleValue.toString().trim()))
    {
      
    if((EnvParams.Country.toUpperCase()=="CHINA")&&(Language=="Chinese (Simplified)")){
        var atSize = JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path, Language, "Client No").OleValue.toString().trim();
      pdflineSplit[j] = pdflineSplit[j].substring(atSize.length+1); 
      x= pdflineSplit[j].split(" ");
      x[0]= pdflineSplit[j];
      x[1]= pdflineSplit[j];
      }else
      x= pdflineSplit[j].split(":");
      pdfJobName = x[1].trim();
        if(pdflineSplit[j].includes(clientNumber))
         {
          ReportUtils.logStep("INFO",clientNumber+"Client Number is matching with Pdf")
          ValidationUtils.verify(true,true,clientNumber+" Client Number is matching with Pdf")
          TextUtils.writeLog(clientNumber+" Client Number is matching with Pdf")
          if(EnvParams.Country.toUpperCase()=="INDIA")
          break;
          }
          else{
          ValidationUtils.verify(false,true,"Client Number is not same in CreditMemo");
          if(EnvParams.Country.toUpperCase()=="INDIA")
           break;
        }
    }

    
//var TaxVariable = "";
//   if((EnvParams.Country.toUpperCase()=="SPAIN") || (EnvParams.Country.toUpperCase()=="MALAYSIA") || (EnvParams.Country.toUpperCase()=="CHINA"))
//   var TaxVariable = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Tax No").OleValue.toString().trim());
//   else if(EnvParams.Country.toUpperCase()=="SINGAPORE")
//   var TaxVariable = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "GST No").OleValue.toString().trim());
////   Log.Message(TaxVariable)
//    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, TaxVariable).OleValue.toString().trim()))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobName = x[1].trim();
//        if(pdflineSplit[j].includes(TaxNo))
//         {
//          ReportUtils.logStep("INFO","Tax No is matching with Pdf")
//          ValidationUtils.verify(true,true,TaxNo+" Tax No is matching with Pdf")
//          TextUtils.writeLog(TaxNo+" Tax No is matching with Pdf")
//          }
//          else{
//          ValidationUtils.verify(false,true,"Tax No is not same in Invoice");
//        }
//    }


//    if(!pName){
////      Log.Message(pName)
//         if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Product:").OleValue.toString().trim()))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobName = x[1].trim();
//        if(pdflineSplit[j].includes(productName))
//         {
//          ReportUtils.logStep("INFO",productName+" Product Name is matching with Pdf")
//          ValidationUtils.verify(true,true,productName+" Product Name is matching with Pdf")
//          TextUtils.writeLog(productName+" Product Name is matching with Pdf")
//          }
//          else{
//          ValidationUtils.verify(false,true,"Product Name is not same in Invoice");
//        }
//    }
//    }
    
  }
  
   if(EnvParams.Country.toUpperCase()=="INDIA"){
   
    var clientGST = ReadExcelSheet("Tax No.",EnvParams.Opco,"CreateClient");
    var pos = ReadExcelSheet("State Code",EnvParams.Opco,"CreateClient");
    var pdfClientGST, pdfPOS;
          
    pointer = pdflineSplit.indexOf("Client GST Details")+1;  // Start searching for client GST details from this Section
       if(pointer>=0){  
           for (j=pointer; j<pdflineSplit.length; j++)
          {
             if(pdflineSplit[j].includes("GSTIN"))
              {
                x= pdflineSplit[j].split(":");
                pdfClientGST = x[1].trim();
               if(clientGST!=pdfClientGST)
                ValidationUtils.verify(false,true,"clientGST is not same in CreditMemo");
               else
               {
                ReportUtils.logStep("INFO",clientGST+" clientGST is matching with Pdf")
                ValidationUtils.verify(true,true,clientGST+" clientGST is matching with Pdf")
                TextUtils.writeLog(clientGST+" clientGST is matching with Pdf")
               }
             }
             if(pdflineSplit[j].includes("Place of Supply"))
              {
                x= pdflineSplit[j].split(":");
                pdfPOS = x[1].trim();
               if(pdfPOS.includes(pos))
               {
                ReportUtils.logStep("INFO",pos+" POS is matching with Pdf")
                ValidationUtils.verify(true,true,pos+" POS is matching with Pdf")
                TextUtils.writeLog(pos+" POS is matching with Pdf")
                break;
                }
               else
               {
                ValidationUtils.verify(false,true,"POS is not same in CreditMemo");
                break;
               }
            }
       } 
      }  
    pointer =-1;  // Setting again pointer to 1
    pointer = pdflineSplit.indexOf("Agency GST Details")+1;
    if(pointer>=0){
    var pdfPan, pdfGstin, pdfCin;  
    var pan = ReadExcelSheet("OpCo PAN",EnvParams.Opco,"OpCo Details");
    var gstin = ReadExcelSheet("OpCo GSTIN",EnvParams.Opco,"OpCo Details");
    var cin = ReadExcelSheet("CIN/UIN",EnvParams.Opco,"OpCo Details");
   
      for (j=pointer; j<pdflineSplit.length; j++)
      {
      if(pdflineSplit[j].includes("PAN"))
      {
        x= pdflineSplit[j].split(":");
        pdfPan = x[1].trim();
         if(pan!=pdfPan)
          ValidationUtils.verify(false,true,"PAN is not same in CreditMemo");
         else{
          ReportUtils.logStep("INFO",pan+" PAN is matching with Pdf")
          ValidationUtils.verify(true,true,pan+" PAN is matching with Pdf")
          TextUtils.writeLog(pan+" PAN is matching with Pdf")
          }
      }
       if(pdflineSplit[j].includes("GSTIN"))
      {
        x= pdflineSplit[j].split(":");
        pdfGstin = x[1].trim();
        if(gstin!=pdfGstin)
            ValidationUtils.verify(false,true,"GSTIN is not same in CreditMemo");
        else{
            ReportUtils.logStep("INFO",gstin+" GSTIN is matching with Pdf")
          ValidationUtils.verify(true,true,gstin+" GSTIN is matching with Pdf")
          TextUtils.writeLog(gstin+" GSTIN is matching with Pdf")
          }
      }
      if(pdflineSplit[j].includes("CIN/UIN"))
      {
        x= pdflineSplit[j].split(":");
        pdfCin = x[1].trim();
        if(cin!=pdfCin)
            ValidationUtils.verify(false,true,"CIN/UIN is not same in CreditMemo");
        else{
            ReportUtils.logStep("INFO",cin+" CIN/UIN is matching with Pdf")
          ValidationUtils.verify(true,true,cin+" CIN/UIN is matching with Pdf")
          TextUtils.writeLog(cin+" CIN/UIN is matching with Pdf")
          }
      }
    }
   }
   else
    ValidationUtils.verify(false,true,"Agency GST Details Section is not displayed in CreditMemo");
  }
  
  
  ExcelUtils.setExcelName(workBook, sheetName, true);
  Log.Message(workBook)
  Log.Message(sheetName)
  for(var i=1;i<11;i++){
  var temp = "";
  var Q_Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco);
  Log.Message(Q_Desp)
  if(Q_Desp!=""){
//  Q_Desp = Q_Desp.replace(/(?![\x00-\x7F])./g, '');
  temp = temp + Q_Desp+" ";
  
  var Q_Qty = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco);
  if(Q_Qty!=""){
  temp = temp + Q_Qty+" ";
  }
  var Q_Billing = ExcelUtils.getColumnDatas("UnitPrice_"+i,EnvParams.Opco);
  if(Q_Billing!="")
    temp = temp + Q_Billing+" ";
  
  var Q_BillingTotal = ExcelUtils.getColumnDatas("TotalBilling_"+i,EnvParams.Opco);
  if(Q_BillingTotal!="")
    temp = temp + Q_BillingTotal+" ";
  Log.Message(EnvParams.Country.toUpperCase())
   if(EnvParams.Country.toUpperCase()=="INDIA")
   { 
        var Q_Tax1 = ExcelUtils.getColumnDatas("Tax1_"+i,EnvParams.Opco);
         var matches = Q_Tax1.match(/(\d+)/); 
         if (matches) 
          temp = temp + matches[1]+".00 ";  
     
        var Q_Tax2 = ExcelUtils.getColumnDatas("Tax2_"+i,EnvParams.Opco);
        if(Q_Tax2!=""){
//           temp = temp + Q_Tax2+" ";
         var matches = Q_Tax2.match(/(\d+)/); 
         if (matches) 
          temp = temp + matches[1]+".00 "; 
          }
  
        var Q_Tax1currency = ExcelUtils.getColumnDatas("Tax1currency_"+i,EnvParams.Opco);
        if(Q_Tax1currency!="")
          if(Q_Tax1currency!="0.00")
           temp = temp + Q_Tax1currency+" ";
  
        var Q_Tax2currency = ExcelUtils.getColumnDatas("Tax2currency_"+i,EnvParams.Opco);
        if(Q_Tax2currency!="")
         if(Q_Tax2currency!="0.00")
           temp = temp + Q_Tax2currency+" ";
 
        var Q_total = ExcelUtils.getColumnDatas("Total_"+i,EnvParams.Opco);
        if(Q_total!=""){
          Q_total = formatMoney(Q_total);
           temp = temp + Q_total;
           }
   }
  
  Log.Message("From Excel :"+temp.trim()) 
  var found = false;
  temp = temp.trim();
   for (z=0; z<pdflineSplit.length; z++){
//     Log.Message(temp.trim()+":"+temp.trim().length)
//     Log.Message(pdflineSplit[z]+":"+pdflineSplit[z].length)
//     Log.Message(" ");
//     Log.Message(pdflineSplit[z].includes(temp.trim()))
//     TextUtils.writeLog("Temp:"+temp.trim()+": Matched with pdf");
//     TextUtils.writeLog("pdflineSplit[z]:"+pdflineSplit[z]+": Matched with pdf");
      if(pdflineSplit[z].includes(temp.trim())){
         ReportUtils.logStep("INFO",temp+" is matching with Pdf")
          ValidationUtils.verify(true,true,temp+" Matched with pdf"); 
          TextUtils.writeLog(temp+" Matched with pdf"); 
        found = true;
        break;
      }
      if(z==pdflineSplit.length-1 && !found){
        ValidationUtils.verify(false,true,temp+" is not matching with the Pdf"); 
        break;
      }
   } 
   }
   else
    break;
  }
  
     var index = docObj.indexOf(JobCurrency+" "+GrandTotal);
    if(index>=0){
          ReportUtils.logStep("INFO",GrandTotal+" Total is matching with Pdf");
          ValidationUtils.verify(true,true,GrandTotal+" Total is matching with Pdf")
          TextUtils.writeLog(GrandTotal+" Total is matching with Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Total is not same in CreditMemo");
}


//function validateInvoicePdf()
//{
//  var fileName = "C:\\Users\\516188\\Documents\\2011\\Print Invoice Editing-6.pdf";
//  var docObj;
//
//  // Load the PDF file to the PDDocument object
//  try{
//  Log.Message(fileName)
//  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
//  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
//  }catch(objEx){
//    Log.Error("Exception while reading document::"+objEx);
//  }
//  var workBook = "C:\\Users\\516188\\Documents\\2011\\DS_MLY_REGRESSION.xlsx";
//  var sheetName = "InvoicePreparation";
//  var country = "Malaysia";
//  EnvParams.Opco = "2011";
//  ExcelUtils.setExcelName(workBook, "Data Management", true);
// 
//  var pdflineSplit = docObj.split("\r\n");
//  
//  var clientName = ReadExcelSheet("Client Name",EnvParams.Opco,"CreateClient");
//  var street = ReadExcelSheet("Street 1",EnvParams.Opco,"CreateClient");
//  var postCode = ReadExcelSheet("Post Code",EnvParams.Opco,"CreateClient");
//  var postDistrict = ReadExcelSheet("Post District",EnvParams.Opco,"CreateClient");
//  var clientCountry = ReadExcelSheet("Country",EnvParams.Opco,"CreateClient");
//  
//  var clientFound = false;
//  for (var j=0; j<pdflineSplit.length; j++)
//  {
//         if(pdflineSplit[j].includes(clientName))
//             {
//             Log.Message(clientName+" clientName is matching with Pdf");
//             clientFound = true;
//             break;
//             }
//         if(j==pdflineSplit.length-1 && !clientFound)
//          ValidationUtils.verify(false,true,"clientName is not found in CreditMemo");
//    
//  }          
//   var index = pdflineSplit.indexOf(street);
//    if(index>=0)
//          Log.Message(street+" Street is matching with Pdf")
//          else
//          ValidationUtils.verify(false,true,"Street is not same in CreditMemo");
//   var index = pdflineSplit.indexOf(postCode+"  "+postDistrict);
//    if (index == -1)
//        index = pdflineSplit.indexOf(postCode+" "+postDistrict);
//    if(index>=0)
//          Log.Message(postCode+" "+postDistrict+" PostCode and Post District are matching with Pdf")
//          else
//          ValidationUtils.verify(false,true,"PostCode and Post District are not same in Invoice");
//   var index = pdflineSplit.indexOf(clientCountry);
//    if(index>=0)
//          Log.Message(clientCountry+" Country is matching with Pdf")
//          else
//          ValidationUtils.verify(false,true,"Country is not same in CreditMemo");
//   
//  jobNumber = ReadExcelSheet("Job Number",EnvParams.Opco,"Data Management");
//  if((jobNumber=="")||(jobNumber==null)){
//  ExcelUtils.setExcelName(workBook, sheetName, true);
//  jobNumber = ExcelUtils.getColumnDatas("Job Number",EnvParams.Opco)
//  }
//  if((jobNumber=="")||(jobNumber==null))
//  ValidationUtils.verify(false,true,"Job Number is needed to Validate CreditMemo");
//   
//  var j, x, pdfJobNum, pointer, pdfJobName;
//  
//  var jobName = ReadExcelSheet("Job_name",EnvParams.Opco,"JobCreation");
//  
//  for (j=0; j<pdflineSplit.length; j++)
//  {
//    if(pdflineSplit[j].includes("Job No"))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobNum = x[1].trim();
//       if(jobNumber!=pdfJobNum)
//        ValidationUtils.verify(false,true,"Job Number is not same in CreditMemo");
//        else
//        Log.Message(jobNumber+" Job Number is matching with Pdf")
//    }
//     if(pdflineSplit[j].includes("Job Name"))
//    {
//      x= pdflineSplit[j].split(":");
//      pdfJobName = x[1].trim();
//      if(pdfJobName.includes(jobName))
//          Log.Message(jobName+" Job Name is matching with Pdf")
//          else
//          ValidationUtils.verify(false,true,"Job Name is not same in CreditMemo");
//      break;
//    }
//  }
//  
//   if(country=="India"){
//    pointer = pdflineSplit.indexOf("Agency GST Details")+1;
//    if(pointer>=0){
//    var pdfPan, pdfGstin, pdfCin;  
//    var pan = ReadExcelSheet("OpCo PAN",EnvParams.Opco,"Data Management");
//    var gstin = ReadExcelSheet("OpCo GSTIN",EnvParams.Opco,"Data Management");
//    var cin = ReadExcelSheet("CIN/UIN",EnvParams.Opco,"Data Management");
//   
//      for (j=pointer; j<pdflineSplit.length; j++)
//      {
//      if(pdflineSplit[j].includes("PAN"))
//      {
//        x= pdflineSplit[j].split(":");
//        pdfPan = x[1].trim();
//         if(pan!=pdfPan)
//          ValidationUtils.verify(false,true,"PAN is not same in CreditMemo");
//         else
//          Log.Message(pan+" PAN is matching with Pdf")
//      }
//       if(pdflineSplit[j].includes("GSTIN"))
//      {
//        x= pdflineSplit[j].split(":");
//        pdfGstin = x[1].trim();
//        if(gstin!=pdfGstin)
//            ValidationUtils.verify(false,true,"GSTIN is not same in CreditMemo");
//        else
//            Log.Message(gstin+" GSTIN is matching with Pdf")
//      }
//      if(pdflineSplit[j].includes("CIN/UIN"))
//      {
//        x= pdflineSplit[j].split(":");
//        pdfCin = x[1].trim();
//        if(cin!=pdfCin)
//            ValidationUtils.verify(false,true,"CIN/UIN is not same in CreditMemo");
//        else
//            Log.Message(cin+" CIN/UIN is matching with Pdf")
//      }
//    }
//   }
//  }
//  
//  ExcelUtils.setExcelName(workBook, sheetName, true);
//  for(var i=1;i<11;i++){
//  var temp = "";
//  var Q_Desp = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco);
//  if(Q_Desp!=""){
//    if(Q_Desp.includes("HSN"))
//    temp = Q_Desp;
//    else
//    {
//      temp = temp + Q_Desp+" ";
//  
//      var Q_Qty = ExcelUtils.getColumnDatas("Quantity_"+i,EnvParams.Opco);
//      if(Q_Qty!=""){
//      temp = temp + Q_Qty+" ";
//      }
//      var Q_Billing = ExcelUtils.getColumnDatas("UnitPrice_"+i,EnvParams.Opco);
//      if(Q_Billing!="")
//        temp = temp + Q_Billing+" ";
//  
//      var Q_BillingTotal = ExcelUtils.getColumnDatas("TotalBilling_"+i,EnvParams.Opco);
//      if(Q_BillingTotal!="")
//        temp = temp + Q_BillingTotal+" ";
//  
//       if(country=="India")
//       { 
//            var Q_Tax1 = ExcelUtils.getColumnDatas("Tax1_"+i,EnvParams.Opco);
//             var matches = Q_Tax1.match(/(\d+)/); 
//             if (matches) 
//              temp = temp + matches[1]+".00 ";  
//     
//            var Q_Tax2 = ExcelUtils.getColumnDatas("Tax2_"+i,EnvParams.Opco);
//            if(Q_Tax2!="")
//               temp = temp + Q_Tax2+" ";
//  
//            var Q_Tax1currency = ExcelUtils.getColumnDatas("Tax1currency_"+i,EnvParams.Opco);
//            if(Q_Tax1currency!="")
//              if(Q_Tax1currency!="0.00")
//               temp = temp + Q_Tax1currency+" ";
//  
//            var Q_Tax2currency = ExcelUtils.getColumnDatas("Tax2currency_"+i,EnvParams.Opco);
//            if(Q_Tax2currency!="")
//             if(Q_Tax2currency!="0.00")
//               temp = temp + Q_Tax2currency+" ";
// 
//            var Q_total = ExcelUtils.getColumnDatas("Total_"+i,EnvParams.Opco);
//            if(Q_total!="")
//               temp = temp + Q_total;
//       }
//  }
//  Log.Message("From Excel :"+temp.trim()) 
//  var found = false;
//   for (z=0; z<pdflineSplit.length; z++){
//      if(pdflineSplit[z].includes(temp.trim())){
//        ValidationUtils.verify(true,true,temp+" is matching with the Pdf");
//         Log.Message(temp+" Matched with pdf"); 
//        found = true;
//        break;
//      }
//      if(z==pdflineSplit.length-1 && !found){
//        ValidationUtils.verify(false,true,temp+" is not matching with the Pdf"); 
//        break;
//      }
//   } 
//   }
//   else
//    break;
//  }
//}



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

function formatMoney(amount, decimalCount = 2, decimal = ".", thousands = ",") {
  try {
    decimalCount = Math.abs(decimalCount);
    decimalCount = isNaN(decimalCount) ? 2 : decimalCount;

    const negativeSign = amount < 0 ? "-" : "";

    let i = parseInt(amount = Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
    let j = (i.length > 3) ? i.length % 3 : 0;

    return negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");
  } catch (e) {
    console.log(e)
  }
};


