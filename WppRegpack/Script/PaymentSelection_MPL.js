//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreatePaymentSelection";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var Duedate="";
var VendorNo="";
var Paymentagent="";
var Paymodemode="";
var ExchangeDate="";
var layoutTypes="";
var Invoicenumber="";
var amount ="";
var fileName = "";

//getting data from datasheet
function getDetails(){
  

  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("PaymentSelectionMpl",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"PaymentSelectionMpl is needed to validate");
  }

ExcelUtils.setExcelName(workBook, "Data Management", true);
Paymentagent = ExcelUtils.getRowDatas("Payment Agent",EnvParams.Opco)
  if((Paymentagent==null)||(Paymentagent=="")){ 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    Paymentagent = ExcelUtils.getRowDatas("Payment_Agent",EnvParams.Opco)
 }
Log.Message(Paymentagent)
if((Paymentagent==null)||(Paymentagent=="")){ 
ValidationUtils.verify(false,true,"Payment Agent is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
Paymodemode = ExcelUtils.getRowDatas("Vendor Invoice Payment Mode",EnvParams.Opco)
  if((Paymodemode==null)||(Paymodemode==""))  { 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    Paymodemode = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco)
   }
Log.Message(Paymodemode)
if((Paymodemode==null)||(Paymodemode=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment Selection");
}


ExcelUtils.setExcelName(workBook, "Data Management", true);
Duedate = ExcelUtils.getRowDatas("Vendor Invoice Due Date",EnvParams.Opco)
  if((Duedate==null)||(Duedate==""))  { 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    Duedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
  }
Log.Message(Duedate)
if((Duedate==null)||(Duedate=="")){ 
ValidationUtils.verify(false,true,"Due Date is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
amount = ExcelUtils.getRowDatas("VendorInvoice Amount",EnvParams.Opco)
  if((amount==null)||(amount==""))  { 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    amount = ExcelUtils.getRowDatas("Amount",EnvParams.Opco)
    }
Log.Message(amount)
if((amount==null)||(amount=="")){ 
ValidationUtils.verify(false,true,"Amount is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
layoutTypes = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(layoutTypes)
if((layoutTypes==null)||(layoutTypes=="")){ 
ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, sheetName, true);
exchangeDate = ExcelUtils.getRowDatas("Exchange_Date",EnvParams.Opco)
Log.Message(exchangeDate)
if((exchangeDate==null)||(exchangeDate=="")){ 
ValidationUtils.verify(false,true,"exchangeDate is Needed to validate a Payment Selection");
}


ExcelUtils.setExcelName(workBook, "Data Management", true);
Invoicenumber = ExcelUtils.getRowDatas("Vendor Invoice NO",EnvParams.Opco)
  if((Invoicenumber==null)||(Invoicenumber==""))  { 
    ExcelUtils.setExcelName(workBook, sheetName, true);
    Invoicenumber = ExcelUtils.getRowDatas("Vendor Invoice NO",EnvParams.Opco)
    }
Log.Message(Invoicenumber)
if((Invoicenumber==null)||(Invoicenumber=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice Number is Needed to Create a Payment Selection");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
Log.Message(VendorNo)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
}
}


function validateCreateChangePaymentSelection_standardLayout()
{
  
  var docObj;

  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
 
  var pdflineSplit = docObj.split("\r\n");
 
  ExcelUtils.setExcelName(workBook, sheetName, true);
                    
  verifyVendorNumber(VendorNo, pdflineSplit);     
  //verifyPaymentAgent(Paymentagent, pdflineSplit);    
  verifyPaymodeMode(Paymodemode,pdflineSplit);          
  //verifyExchangeDate(exchangeDate,pdflineSplit);
  verifyDueDate(Duedate,pdflineSplit);     
  verifyAmount(amount,pdflineSplit);
 }


function validateCreateChangePaymentSelection_wppLayout()
{
  

  var docObj;
  
  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
 
  var pdflineSplit = docObj.split("\r\n");
               
  verifyVendorNumber(VendorNo, pdflineSplit);
  verifyInvoiceNumber(Invoicenumber,pdflineSplit);          
  verifyAmount(amount,pdflineSplit);
 // verifyExchangeDate(exchangeDate,pdflineSplit);
  verifyDueDate(Duedate,pdflineSplit);     
  verifyPaymodeMode(Paymodemode,pdflineSplit);  
          
}



//Main Function
function validatePaymentSelectionMPL() {
TextUtils.writeLog("Create Payment Selection Started"); 

Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

try{
  
getDetails();
if(layoutTypes=="WPP Payment")
{
validateCreateChangePaymentSelection_wppLayout();
}
else if(layoutTypes=="Standard")
{
  validateCreateChangePaymentSelection_standardLayout();
}
 
}
  catch(err){
    Log.Message(err);
  }

}


function verifyVendorNumber(vendorNumber,pdflineSplit)
{
    var vendorNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(vendorNumber))
             {
             Log.Message(vendorNumber+" vendorNumber is matching with Pdf");
             ValidationUtils.verify(true,true,"VendorNumber is matched in with Pdf");
             vendorNoFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !vendorNoFound)
          ValidationUtils.verify(false,true,"VendorNumber is not same in CreatePaymentSelection");
  }  
}

function verifyInvoiceNumber(vendorInvoiceNo,pdflineSplit)
{
  var vendorInvoiceNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
          if(vendorInvoiceNo.includes(pdflineSplit[j]))             {
             Log.Message(vendorInvoiceNo+" vendorInvoiceNo is matching with Pdf");
             ValidationUtils.verify(true,true,"Vendor Invoice No is matched with Pdf");
             vendorInvoiceNoFound = true;
             break;
             }
         else
         continue;
         if(j==pdflineSplit.length-1 && !vendorInvoiceNoFound)
          ValidationUtils.verify(false,true,"vendorInvoiceNo is not same in CreatePaymentSelection");
    
  }       
}

function verifyAmount(amount,pdflineSplit)
{
  var amountFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(amount))
             {
             Log.Message(amount+" amount is matching with Pdf");
             ValidationUtils.verify(true,true,"Amount is matched with Pdf");
             amountFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !amountFound)
          ValidationUtils.verify(false,true,"amount is not same in CreatePaymentSelection");
    
    }
}

function verifyExchangeDate(exchangeDate,pdflineSplit)
{
  var exchangeDateFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(exchangeDate))
             {
             Log.Message(exchangeDate+" exchangeDate is matching with Pdf");
             ValidationUtils.verify(true,true,"Exchange Date is matched with Pdf");
             exchangeDateFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !exchangeDateFound)
          ValidationUtils.verify(false,true,"exchangeDate is not same in CreatePaymentSelection");
    
    } 
}

function verifyDueDate(dueDate,pdflineSplit)
{
     var dueDateFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(dueDate))
             {
             Log.Message(dueDate+" DueDate is matching with Pdf");
             ValidationUtils.verify(true,true,"Due Date is matched with Pdf");
             dueDateFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !dueDateFound)
          ValidationUtils.verify(false,true,"DueDate is not same in CreatePaymentSelection");
    
    }    
}

function verifyPaymodeMode(paymodeMode, pdflineSplit)
{
   var paymodeModeFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(paymodeMode))
             {
             Log.Message(paymodeMode+" paymodeMode is matching with Pdf");
             ValidationUtils.verify(true,true,"Payment Mode is matched with Pdf");
             paymodeModeFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymodeModeFound)
          ValidationUtils.verify(false,true,"paymodeMode is not same in CreatePaymentSelection");    
    }
}
function verifyPaymentAgent(paymentAgent,pdflineSplit)
{
   var paymentAgentFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(paymentAgent))
             {
             Log.Message(paymentAgent+" paymentAgent is matching with Pdf");
             ValidationUtils.verify(true,true,"Payment Agent is matched with Pdf");
             paymentAgentFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymentAgentFound)
          ValidationUtils.verify(false,true,"paymentAgent is not same in CreatePaymentSelection");    
    }
}



