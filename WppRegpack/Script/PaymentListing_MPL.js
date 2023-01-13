//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "CreatePaymentFile";
var Language = "";
Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var dueDate="";
var VendorNo="";
var paymentAgent="";
var paymentMode="";
var ExchangeDate="";
var PrintLayout="";
var Invoicenumber="";
var PaymentDate="";
var PaymentNo="";
var filepathforMplValidation ="";
var fileName = "";
var amount ="";
var docObj = "";


//getting data from datasheet
function getDetails(){

ExcelUtils.setExcelName(workBook, "Data Management", true);
fileName = ExcelUtils.getRowDatas("PDF Payment File",EnvParams.Opco)
if((fileName==null)||(fileName=="")){ 
 ValidationUtils.verify(false,true,"PDF Payment File is needed to Validate Payment Listing");
}

paymentAgent = ExcelUtils.getRowDatas("Payment Agent",EnvParams.Opco)
Log.Message(paymentAgent)
if((paymentAgent==null)||(paymentAgent=="")){ 
ValidationUtils.verify(false,true,"Payment Agent is Needed to Validate Payment Listing");
}

paymentMode = ExcelUtils.getRowDatas("Vendor Invoice Payment Mode",EnvParams.Opco)
Log.Message(paymentMode)
if((paymentMode==null)||(paymentMode=="")){ 
ValidationUtils.verify(false,true,"paymentMode is Needed to Validate Payment Listing");
}


PaymentNo = ExcelUtils.getRowDatas("Payment Number",EnvParams.Opco)
Log.Message(PaymentNo)
if((PaymentNo==null)||(PaymentNo=="")){ 
ValidationUtils.verify(false,true,"PaymentNumber is Needed to Validate Payment Listing");
}

Invoicenumber = ExcelUtils.getRowDatas("Vendor Invoice NO",EnvParams.Opco)
Log.Message(Invoicenumber)
if((Invoicenumber==null)||(Invoicenumber=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice Nunber is Needed to Validate Payment Listing");
}

dueDate = ExcelUtils.getRowDatas("Vendor Invoice Due Date",EnvParams.Opco)
Log.Message(dueDate)
if((dueDate==null)||(dueDate=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice DueDate is Needed to Validate Payment Listing");
}

amount = ExcelUtils.getRowDatas("VendorInvoice Amount",EnvParams.Opco)
Log.Message(amount)
if((amount==null)||(amount=="")){ 
ValidationUtils.verify(false,true,"Amount is Needed to Validate Payment Listing");
}

ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
Log.Message(VendorNo)
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Validate Payment Listing");
}


}


function validatePaymentListing()
{
  
  var pdflineSplit = docObj.split("\r\n");
                  
  verifyVendorNumber(VendorNo, pdflineSplit);     
  verifyInvoiceNumber(Invoicenumber,pdflineSplit);  
  verifyAmount(amount,pdflineSplit);
  
 }



//Main Function
function paymentListing() {
TextUtils.writeLog("Payment Listing MPL validation Started"); 
Indicator.PushText("waiting for window to open");
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

  getDetails();

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }

validatePaymentListing();
}


function verifyVendorNumber(vendorNumber,pdflineSplit)
{
    var vendorNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(vendorNumber))
             {
             Log.Message(vendorNumber+" vendorNumber is matching with Pdf");
             vendorNoFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !vendorNoFound)
          ValidationUtils.verify(false,true,"VendorNumber is not same in Payment Listing");
  }  
}

function verifyInvoiceNumber(vendorInvoiceNo,pdflineSplit)
{
  var vendorInvoiceNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
          if(vendorInvoiceNo.includes(pdflineSplit[j]))             {
             Log.Message(vendorInvoiceNo+" vendorInvoiceNo is matching with Pdf");
             vendorInvoiceNoFound = true;
             break;
             }
         else
         continue;
         if(j==pdflineSplit.length-1 && !vendorInvoiceNoFound)
          ValidationUtils.verify(false,true,"vendorInvoiceNo is not same in Payment Listing");
    
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
             amountFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !amountFound)
          ValidationUtils.verify(false,true,"amount is not same in Payment Listing");
    
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
             exchangeDateFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !exchangeDateFound)
          ValidationUtils.verify(false,true,"exchangeDate is not same in Payment Listing");
    
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
             dueDateFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !dueDateFound)
          ValidationUtils.verify(false,true,"DueDate is not same in Payment Listing");
    
    }    
}
function verifyPaymentNumber(paymentNumber,pdflineSplit)
{
   var paymentNumberFound = false;
    for (var j=0; j<pdflineSplit.length; j++)
    {
         if(pdflineSplit[j].includes(paymentNumber))
             {
             Log.Message(paymentNumber+" PaymentNumber is matching with Pdf");
             paymentNumberFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymentNumberFound)
          ValidationUtils.verify(false,true,"PaymentNumber is not same in Payment Listing");    
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
             paymodeModeFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymodeModeFound)
          ValidationUtils.verify(false,true,"paymodeMode is not same in Payment Order");    
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
             paymentAgentFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !paymentAgentFound)
          ValidationUtils.verify(false,true,"paymentAgent is not same in Payment Order");    
    }
}




