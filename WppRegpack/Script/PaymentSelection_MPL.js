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
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var Duedate="";
var VendorNo="";
var Paymentagent="";
var Paymodemode="";
var ExchangeDate="";
var layoutTypes="";
var Invoicenumber="";
var amount ="";

//getting data from datasheet
function getDetails(){
ExcelUtils.setExcelName(workBook, sheetName, true);
Paymentagent = ExcelUtils.getRowDatas("Payment_Agent",EnvParams.Opco)
Log.Message(Paymentagent)
if((Paymentagent==null)||(Paymentagent=="")){ 
ValidationUtils.verify(false,true,"Payment Agent is Needed to Create a Payment Selection");
}
Paymodemode = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco)
Log.Message(Paymodemode)
if((Paymodemode==null)||(Paymodemode=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment Selection");
}
PrintLayout = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco)
Log.Message(PrintLayout)
if((PrintLayout==null)||(PrintLayout=="")){ 
ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment Selection");
}
Duedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
Log.Message(Duedate)
if((Duedate==null)||(Duedate=="")){ 
ValidationUtils.verify(false,true,"Due Date Number is Needed to Create a Payment Selection");
}
amount = ExcelUtils.getRowDatas("Amount",EnvParams.Opco)
Log.Message(amount)
if((amount==null)||(amount=="")){ 
ValidationUtils.verify(false,true,"Amount is Needed to Create a Payment Selection");
}
layoutTypes = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(layoutTypes)
if((layoutTypes==null)||(layoutTypes=="")){ 
ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Selection");
}
Invoicenumber = ExcelUtils.getRowDatas("Vendor Invoice NO",EnvParams.Opco)
Log.Message(Invoicenumber)
if((Invoicenumber==null)||(Invoicenumber=="")){ 
ValidationUtils.verify(false,true,"Vendor Invoice Nunber is Needed to Create a Payment Selection");
}
ExcelUtils.setExcelName(workBook, "Data Management", true);
VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
Log.Message(VendorNo)
if((VendorNo=="")||(VendorNo==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
Log.Message(VendorNo)
}
if((VendorNo==null)||(VendorNo=="")){ 
ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Selection");
}
}





function validateCreateChangePaymentSelection_standardLayout(workBook,sheetName)
{
  
  var fileName = "";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("PaymentSelectionMpl",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"PaymentSelectionMpl is needed to validate");
  }
  
 // var fileName = filepathforMplValidation;
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
  var vendorNumber = ReadExcelSheet("Vendor Number",EnvParams.Opco,sheetName);
  var paymentAgent  = ReadExcelSheet("Payment_Agent",EnvParams.Opco,sheetName);
  var paymodeMode = ReadExcelSheet("Paymode_Mode",EnvParams.Opco,sheetName);
  var exchangeDate = ReadExcelSheet("ExchangeRateDate",EnvParams.Opco,sheetName);
  var dueDate = ReadExcelSheet("Latest Due Date",EnvParams.Opco,sheetName);
  var amount= ReadExcelSheet("Amount",EnvParams.Opco,sheetName);
                    
  verifyVendorNumber(vendorNumber, pdflineSplit);     
  verifyPaymentAgent(paymentAgent, pdflineSplit);    
  verifyPaymodeMode(paymodeMode,pdflineSplit);          
  verifyExchangeDate(exchangeDate,pdflineSplit);
  verifyDueDate(dueDate,pdflineSplit);     
  verifyAmount(amount,pdflineSplit);
 }


function validateCreateChangePaymentSelection_wppLayout(workBook,sheetName)
{
    var fileName = "";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("PaymentSelectionMpl",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"PaymentSelectionMpl is needed to validate");
  }
  
  
 // var fileName = filepathforMplValidation;
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
 // var workBook = "C:\\GlobalTestSuiteAutomation_Bank\\WppRegpack\\TestResource\\Regression\\DS_SPN_REGRESSION.xlsx";
 //  var country = "Spain";
  //EnvParams.Opco = "1006";
 
  var pdflineSplit = docObj.split("\r\n");
 
  ExcelUtils.setExcelName(workBook, sheetName, true);
  var vendorNumber = ReadExcelSheet("Vendor Number",EnvParams.Opco,sheetName);
  var vendorInvoiceNo = ReadExcelSheet("Vendor Invoice NO",EnvParams.Opco,sheetName);
  var amount= ReadExcelSheet("Amount",EnvParams.Opco,sheetName);
  var exchangeDate = ReadExcelSheet("ExchangeDate",EnvParams.Opco,sheetName);
  var dueDate = ReadExcelSheet("Due Date",EnvParams.Opco,sheetName);
  var paymodeMode = ReadExcelSheet("Paymode_Mode",EnvParams.Opco,sheetName);
               
  verifyVendorNumber(vendorNumber, pdflineSplit);
  verifyInvoiceNumber(vendorInvoiceNo,pdflineSplit);          
  verifyAmount(amount,pdflineSplit);
  verifyExchangeDate(exchangeDate,pdflineSplit);
  verifyDueDate(dueDate,pdflineSplit);     
  verifyPaymodeMode(paymodeMode,pdflineSplit);  
          
}





//Main Function
function CreatePayment() {
TextUtils.writeLog("Create Payment Selection Started"); 
Indicator.PushText("waiting for window to open");
Language = "";
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;

excelName = EnvParams.path;
workBook = Project.Path+excelName;
sheetName = "CreatePaymentSelection";

ExcelUtils.setExcelName(workBook, sheetName, true);
Arrays = [];
count = true;
checkmark = false;
STIME = "";
VendorNo,Paymentagent,Paymodemode ="";

Log.Message(Language)
STIME = WorkspaceUtils.StartTime();
TextUtils.writeLog("Execution Start Time :"+STIME); 
ReportUtils.logStep("INFO", "Execution Start Time :"+STIME);

try{
var  layoutType = layoutTypes;

if(layoutType=="WPP Payment")
{
validateCreateChangePaymentSelection_wppLayout(filepathforMplValidation,workBook,sheetName)
}
else if(layoutType=="Standard")
{
  validateCreateChangePaymentSelection_standardLayout(filepathforMplValidation,workBook,sheetName)
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
          ValidationUtils.verify(false,true,"VendorNumber is not same in Create Payment Selection");
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
          ValidationUtils.verify(false,true,"vendorInvoiceNo is not same in CreatePaymentFile");
    
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
          ValidationUtils.verify(false,true,"amount is not same in CreatePaymentFile");
    
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
          ValidationUtils.verify(false,true,"exchangeDate is not same in CreatePaymentFile");
    
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
          ValidationUtils.verify(false,true,"DueDate is not same in CreatePaymentFile");
    
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
          ValidationUtils.verify(false,true,"PaymentNumber is not same in PrintReimmittance");    
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
          ValidationUtils.verify(false,true,"paymodeMode is not same in CreatePaymentSelection/ChangePaymentSelection");    
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
          ValidationUtils.verify(false,true,"paymentAgent is not same in CreatePaymentSelection/ChangePaymentSelection");    
    }
}



