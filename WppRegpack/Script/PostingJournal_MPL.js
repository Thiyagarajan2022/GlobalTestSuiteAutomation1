//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT WorkspaceUtils
//USEUNIT ReportUtils
//USEUNIT ValidationUtils

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;


function validatePostingJournal()
{
  //Setting Language in WorkspaceUtils
  Language = "";
  Language = EnvParams.LanChange(EnvParams.Language);
  WorkspaceUtils.Language = Language;  
  
//  var fileName = "C:\\Users\\516188\\Documents\\Posting Journal\\Print Posting Journal-57_P.pdf";
  var fileName = "";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
  fileName = ExcelUtils.getRowDatas("PDF Print General Journal",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"General Journal PDF is needed to validate");
  }
  var docObj;

  // Load the PDF file to the PDDocument object
  try{
  Log.Message(fileName)
  docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(fileName);
  docObj = getTextFromPDF(docObj).OleValue.toString().trim();
  }catch(objEx){
    Log.Error("Exception while reading document::"+objEx);
  }
//  var workBook = "C:\\Users\\516188\\Documents\\Posting Journal\\DS_SPN_REGRESSION - P.xlsx";
  ExcelUtils.setExcelName(workBook, "Data Management", true);
   var sheetName = "GLMPL";
 // EnvParams.Country = "India";
 // EnvParams.Opco = "1707";
  var pdflineSplit = docObj.split("\r\n");

  ExcelUtils.setExcelName(workBook, "CountryCurrency", true);  
  var baseCurrency = ReadExcelSheet(EnvParams.Country,"Currency","CountryCurrency");
  
  ExcelUtils.setExcelName(workBook, "GLMPL", true); 
  var journalNo = ExcelUtils.getColumnDatas("JOURNAL NO",EnvParams.Opco);
  var companyName  = ExcelUtils.getColumnDatas("Company Name",EnvParams.Opco);
  var submittedBy = ExcelUtils.getColumnDatas("Submitted By",EnvParams.Opco);
  var submittedOn = ExcelUtils.getColumnDatas("Submitted On",EnvParams.Opco);
  var periodStart = ExcelUtils.getColumnDatas("Period start",EnvParams.Opco);
  var periodEnd = ExcelUtils.getColumnDatas("Period end",EnvParams.Opco);
  var postedBy = ExcelUtils.getColumnDatas("Posted By",EnvParams.Opco);
  var postedOn = ExcelUtils.getColumnDatas("Posted On",EnvParams.Opco);
  
  for (j=0; j<pdflineSplit.length; j++)
  {
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "JOURNAL NO").OleValue.toString().trim()))
    {
       if(pdflineSplit[j].includes(journalNo))
       Log.Message(journalNo+" Journal Number is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Journal Number is not same in pdf");
    }
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company No").OleValue.toString().trim()))
    {
       if(pdflineSplit[j].includes(EnvParams.Opco))
       Log.Message(EnvParams.Opco+" Company Number is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Company Number is not same in pdf");
    }
    if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company Name").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(companyName))
       Log.Message(companyName+" Company Name is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Company Name is not same in pdf");
      }
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Base Currency").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(baseCurrency))
       Log.Message(baseCurrency+" Base Currency is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Base Currency is not same in pdf");
      }
      
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submitted By").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(submittedBy))
       Log.Message(submittedBy+" Submitted By is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Submitted By is not same in pdf");
      }
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Submitted On").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(submittedOn))
       Log.Message(submittedOn+" Submitted On is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Submitted On is not same in pdf");
      }
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Posted By").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(postedBy))
       Log.Message(postedBy+" Posted By is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Posted By is not same in pdf");
      }
    
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Period Start").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(periodStart))
       Log.Message(periodStart+" Period Start is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Period Start is not same in pdf");
      }
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Period End").OleValue.toString().trim()))
     {
       if(pdflineSplit[j].includes(periodEnd))
        Log.Message(periodEnd+" Period End is matching with Pdf")
        else
        ValidationUtils.verify(false,true,"Period End is not same in pdf");
      }
      if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Posted On").OleValue.toString().trim()))
      {
       if(pdflineSplit[j].includes(postedOn)){
        Log.Message(postedOn+" Posted On is matching with Pdf")
         break;
        }
        else{
        ValidationUtils.verify(false,true,"Posted On is not same in pdf");
        break;
        }
      }    
  }
  
  var transNo, entryDate, localAccount, description, department, businessUnit, credit, debit;
  for(var i=0;i<2;i++){
   transNo = ExcelUtils.getColumnDatas("Transaction No_"+i,EnvParams.Opco);
   entryDate = ExcelUtils.getColumnDatas("Entry Date_"+i,EnvParams.Opco);
   localAccount = ExcelUtils.getColumnDatas("Local Account_"+i,EnvParams.Opco);
   if(i==1)
   localAccount = localAccount.split("-")[0];
   description = ExcelUtils.getColumnDatas("Description_"+i,EnvParams.Opco);
   department = ExcelUtils.getColumnDatas("Department_"+i,EnvParams.Opco);
   businessUnit = ExcelUtils.getColumnDatas("BusinessUnit_"+i,EnvParams.Opco);
   credit = ExcelUtils.getColumnDatas("Credit_"+i,EnvParams.Opco);
   debit = ExcelUtils.getColumnDatas("Debit_"+i,EnvParams.Opco);
  if(transNo!=""){
          for (j=0; j<pdflineSplit.length; j++)
         {
          if(pdflineSplit[j].includes(localAccount))
          {
              Log.Message(localAccount+" Local Account is matching with Pdf")
             
                if(pdflineSplit[j].includes(entryDate))
                    Log.Message(entryDate+" Entry Date is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Entry Date is not same in pdf");
              
                if(pdflineSplit[j].includes(transNo))
                    Log.Message(transNo+" Transaction Number is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Local Account is not same in pdf");
                }
                if(pdflineSplit[j].includes(description))
                {
                    Log.Message(description+" Description is matching with Pdf")
                 
                if(pdflineSplit[j].includes(department))
                    Log.Message(department+" Department is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Department is not same in pdf");   
           
                if(pdflineSplit[j].includes(businessUnit))
                    Log.Message(businessUnit+" Business Unit is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Business Unit is not same in pdf"); 
            
                if(pdflineSplit[j].includes(credit))
                    Log.Message(credit+" Credit is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Credit is not same in pdf"); 
            
                if(pdflineSplit[j].includes(debit))
                    Log.Message(debit+" Debit is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Debit is not same in pdf");  
                  break;                        
               }
              else if(j==pdflineSplit.length-1)
              ValidationUtils.verify(false,true,"Transaction Number is not same in pdf");
          }
    }
 }
  
 var pointer = -1, endPointer=-1, jobNo, jobType;
 transNo = ExcelUtils.getColumnDatas("Transaction No_1",EnvParams.Opco);
 entryDate = ExcelUtils.getColumnDatas("Entry Date_1",EnvParams.Opco);
 jobNo = ExcelUtils.getColumnDatas("Job No_1",EnvParams.Opco);
 localAccount = ExcelUtils.getColumnDatas("Local Account_1",EnvParams.Opco);
 localAccount = localAccount.split("-")[0];
 description = ExcelUtils.getColumnDatas("Description_1",EnvParams.Opco);
 department = ExcelUtils.getColumnDatas("Department_1",EnvParams.Opco);
 businessUnit = ExcelUtils.getColumnDatas("BusinessUnit_1",EnvParams.Opco);
 cost = ExcelUtils.getColumnDatas("Debit_1",EnvParams.Opco);
 jobType = ExcelUtils.getColumnDatas("Job Type_1",EnvParams.Opco);
   
 pointer = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Job").OleValue.toString().trim())+1;  // Start searching from this Section
 
 endPointer = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Accounts Receivable").OleValue.toString().trim());  // Start searching from this Section
 if (endPointer<0)
 endPointer = pdflineSplit.indexOf(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Accounts Payable").OleValue.toString().trim());  // Start searching from this Section
 
       if(pointer>=0){  
           for (j=pointer; j<endPointer; j++)
          {
             if(pdflineSplit[j].includes(transNo))
              {
               if(pdflineSplit[j].includes(entryDate))
                    Log.Message(entryDate+" Entry Date is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Entry Date is not same in pdf");
             
                 if(pdflineSplit[j].includes(jobNo))
                    Log.Message(jobNo+" Job No is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Job No is not same in pdf");
                 } 
               if(pdflineSplit[j].includes(description))
              {
                Log.Message(description+" Description is matching with Pdf")        
                if(pdflineSplit[j].includes(department))
                    Log.Message(department+" Department is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Department is not same in pdf");   
           
                if(pdflineSplit[j].includes(businessUnit))
                    Log.Message(businessUnit+" Business Unit is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Business Unit is not same in pdf"); 
            
                if(pdflineSplit[j].includes(cost))
                    Log.Message(cost+" cost is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"cost is not same in pdf");  
                  break;                        
                }
              else if(j==pdflineSplit.length-1)
              ValidationUtils.verify(false,true,"Transaction Number is not same in pdf");
            }
       }
 
 var credit,grp,vendorNo,vendorName,vendorCurrency,clientNo,clientName,clientCurrency;
 transNo = ExcelUtils.getColumnDatas("Transaction No_0",EnvParams.Opco);
 entryDate = ExcelUtils.getColumnDatas("Entry Date_0",EnvParams.Opco);
 description = ExcelUtils.getColumnDatas("Description_0",EnvParams.Opco);
 credit = ExcelUtils.getColumnDatas("Credit_0",EnvParams.Opco);
 grp =ExcelUtils.getColumnDatas("GRP_0",EnvParams.Opco); 
 
 ExcelUtils.setExcelName(workBook, "Data Management", true);
 vendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
 vendorName = ReadExcelSheet("Global Vendor Name",EnvParams.Opco,"Data Management");
 vendorCurrency = ReadExcelSheet("Global Vendor Currency",EnvParams.Opco,"Data Management");
 
 clientNo = ReadExcelSheet("Global Client Number",EnvParams.Opco,"Data Management");
 clientName = ReadExcelSheet("Global Client Name",EnvParams.Opco,"Data Management");
 clientCurrency = ReadExcelSheet("Global Client Currency",EnvParams.Opco,"Data Management");

       if(endPointer>=0){  
           for (j=endPointer; j<pdflineSplit.length; j++)
          {
             if(pdflineSplit[j].includes(transNo))
              {
               if(pdflineSplit[j].includes(entryDate))
                    Log.Message(entryDate+" Entry Date is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Entry Date is not same in pdf");
             
                if(pdflineSplit[j].includes(description))
                    Log.Message(description+" Description is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Description is not same in pdf"); 
                    
                   if(pdflineSplit[j].includes(credit))
                    Log.Message(credit+" Credit is matching with Pdf")
                      else
                    ValidationUtils.verify(false,true,"Credit is not same in pdf");        
           
              if(grp=="P")      
                    { 
                        if(pdflineSplit[j].includes(vendorNo))
                        Log.Message(vendorNo+" vendorNumber is matching with Pdf")
                          else
                        ValidationUtils.verify(false,true,"vendorNumber is not same in pdf");   
           
                       if(pdflineSplit[j].includes(vendorName))
                        Log.Message(vendorName+" vendor Name is matching with Pdf")
                          else
                        ValidationUtils.verify(false,true,"vendor Name is not same in pdf"); 

                      if(pdflineSplit[j].includes(vendorCurrency))
                        Log.Message(vendorCurrency+" vendorCurrency is matching with Pdf")
                          else
                        ValidationUtils.verify(false,true,"vendorCurrency is not same in pdf");  
                      break;
                      }
                    if(grp=="R")      
                    { 
                        if(pdflineSplit[j].includes(clientNo))
                        Log.Message(clientNo+" clientNumber is matching with Pdf")
                          else
                        ValidationUtils.verify(false,true,"clientNumber is not same in pdf");   
           
                       if(pdflineSplit[j].includes(clientName))
                        Log.Message(clientName+" clientName is matching with Pdf")
                          else
                        ValidationUtils.verify(false,true,"clientName is not same in pdf"); 

                      if(pdflineSplit[j].includes(clientCurrency))
                        Log.Message(clientCurrency+" clientCurrency is matching with Pdf")
                          else
                        ValidationUtils.verify(false,true,"clientCurrency is not same in pdf");  
                      break;
                      }                     
               }
              else if(j==pdflineSplit.length-1)
              ValidationUtils.verify(false,true,"Transaction Number is not same in pdf");
            }
       }       
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