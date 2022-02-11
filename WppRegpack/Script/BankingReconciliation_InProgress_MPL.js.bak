//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Banking Reconciliation MPL";
var Language = "";
  Indicator.Show();
  
ExcelUtils.setExcelName(workBook, sheetName, true);
var STIME = "";
var bankAccountNumber="";
var account = "";
var localAccountNumber ="";
var reconcilitionDate ="";
var transactionNumber = "";
var jounralNumber = "";
var text = "";
var debit = "";
var credit = "";
var statementDate = "";
var statementNo = "";
var fileName = "";

//getting data from datasheet
function getDetails(){
  
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("Bank_Reconciliation In-Progress PDF",EnvParams.Opco)
  Log.Message(fileName);
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"Bank_Reconciliation PDF is needed to validate");
  }
  
  ExcelUtils.setExcelName(workBook, sheetName, true);
  bankAccountNumber = ExcelUtils.getColumnDatas("Bank Acc. No",EnvParams.Opco)
  Log.Message(bankAccountNumber)
  if((bankAccountNumber==null)||(bankAccountNumber=="")){ 
  ValidationUtils.verify(false,true,"Bank Account Number is Needed to validate Bank Reconciliation");
  }

  account = ExcelUtils.getColumnDatas("Account",EnvParams.Opco)
  Log.Message(account)
  if((account==null)||(account=="")){ 
  ValidationUtils.verify(false,true,"Account is Needed to validate Bank Reconciliation");
  }

  localAccountNumber = ExcelUtils.getColumnDatas("Local Account No",EnvParams.Opco)
  Log.Message(localAccountNumber)
  if((localAccountNumber==null)||(localAccountNumber=="")){ 
  ValidationUtils.verify(false,true,"Local Account Number is Needed to validate Bank Reconciliation");
  }

  reconcilitionDate = ExcelUtils.getColumnDatas("Reconcilition Date",EnvParams.Opco)
  Log.Message(reconcilitionDate)
  if((reconcilitionDate==null)||(reconcilitionDate=="")){ 
  ValidationUtils.verify(false,true,"reconcilitionDate is Needed to validate Bank Reconciliation");
  }

  transactionNumber = ExcelUtils.getColumnDatas("Trans No",EnvParams.Opco)
  Log.Message(transactionNumber)
  if((transactionNumber==null)||(transactionNumber=="")){ 
  ValidationUtils.verify(false,true,"Transaction Number is Needed to validate Bank Reconciliation");
  }

  jounralNumber = ExcelUtils.getColumnDatas("Journal",EnvParams.Opco)
  Log.Message(jounralNumber)
  if((jounralNumber==null)||(jounralNumber=="")){ 
  ValidationUtils.verify(false,true,"jounralNumber is Needed to validate Bank Reconciliation");
  }

  text = ExcelUtils.getColumnDatas("TEXT",EnvParams.Opco)
  Log.Message(text)
  if((text==null)||(text=="")){ 
  ValidationUtils.verify(false,true,"text is Needed to validate Bank Reconciliation");
  }

  debit = ExcelUtils.getColumnDatas("DEBIT",EnvParams.Opco)
  Log.Message(debit)
  if((debit==null)||(debit=="")){ 
  ValidationUtils.verify(false,true,"debit is Needed to validate Bank Reconciliation");
  }

  credit = ExcelUtils.getColumnDatas("CREDIT",EnvParams.Opco)
  Log.Message(credit)
  if((credit==null)||(credit=="")){ 
  ValidationUtils.verify(false,true,"credit is Needed to validate Bank Reconciliation");
  }

}


function Banking_Reconciliation_InProgress()
{
  
  var docObj;
  try{
  getDetails();
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
  
  var pdflineSplit = docObj.split("\r\n");
 
   index = pdflineSplit.indexOf("BANK RECONCILIATION");  
     if(index>=0){
          ReportUtils.logStep("INFO","Heading is available Pdf")
          ValidationUtils.verify(true,true,"Heading is available Pdf")
          TextUtils.writeLog("Heading is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Heading is not available Pdf") 
   
   
     index = docObj.indexOf(EnvParams.Opco);  
     if(index>=0){
          ReportUtils.logStep("INFO","Opco Number is available Pdf")
          ValidationUtils.verify(true,true,"Opco Number is available Pdf")
          TextUtils.writeLog("Opco Number is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Opco Number is not available Pdf")      
 
     index = docObj.indexOf(bankAccountNumber);  
     if(index>=0){
          ReportUtils.logStep("INFO","Bank Account Number is available Pdf")
          ValidationUtils.verify(true,true,"Bank Account Number is available Pdf")
          TextUtils.writeLog("Bank Account Number is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Bank Account Number is not available Pdf")          
     
     index = docObj.indexOf(account);  
     if(index>=0){
          ReportUtils.logStep("INFO","Account Number is available in Pdf")
          ValidationUtils.verify(true,true," Account Number is available in Pdf")
          TextUtils.writeLog(" Account Number is available in Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Bank Account Number is not available in Pdf")  
          
     index = docObj.indexOf(localAccountNumber);  
     if(index>=0){
          ReportUtils.logStep("INFO","Local Account Number is available in Pdf")
          ValidationUtils.verify(true,true,"Local Account Number is available in Pdf")
          TextUtils.writeLog("Local Account Number is available in Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Local Account Number is not available in Pdf")  
          
          
     var statemententry = reconcilitionDate+" "+transactionNumber+" "+jounralNumber+" "+text+" "+debit+" "+credit; 
     Log.Message(statemententry);                  
     index = docObj.includes(statemententry);  
     if(index){
          ReportUtils.logStep("INFO", statemententry+" Statement Entry line is available in Pdf")
          ValidationUtils.verify(true,true,"Statement Entry line is available in Pdf")
          TextUtils.writeLog("Statement Entry line is available Pdf")
          }
          else
          ValidationUtils.verify(false,true,"Statement Entry line is not available in Pdf") 
      
          
      var textStrings =['FOR RECONCILIATION','NOT FOR RECONCILIATION']; 
      
      textStrings.forEach(function(textString){
      index = docObj.includes(textString);
      if(index){
          ReportUtils.logStep("INFO", textString+"  is available in Pdf")
          ValidationUtils.verify(true,true,textString+" is available in Pdf")
          TextUtils.writeLog(textString+" is available in Pdf")
          }
          else
          ValidationUtils.verify(false,true,textString+" is not available in Pdf") ;    
      })
            
          
           
}