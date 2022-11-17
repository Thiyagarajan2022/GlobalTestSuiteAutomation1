//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT PdfUtils
//USEUNIT ReportUtils
//USEUNIT Restart
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils

var excelName = EnvParams.getEnvironment();
var workBook = Project.Path+excelName;
Indicator.Show();
Indicator.PushText("waiting for window to open");

var Language = "";
Language = EnvParams.LanChange(EnvParams.Language);


ExcelUtils.setExcelName(workBook, sheetName, true);
Log.Message(sheetName);
var STIME = "";
var Expense_Number, fileName, pdflineSplit = "";


function MPLExpenses() {

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
ExcelUtils.setExcelName(workBook, "Server Details", true);
var Project_manager = ExcelUtils.getRowDatas("UserName",EnvParams.Opco)
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
}

STIME = WorkspaceUtils.StartTime();
excelName = EnvParams.path;
workBook = Project.Path+excelName

  getDetails();
  goToExpeseMenuItem();
  gotoTimeExpenses();
  print();
  validateExpenseSheet();
  WorkspaceUtils.closeAllWorkspaces();    
}


function getDetails(){
  sheetName = "ExpensesMPL";
ExcelUtils.setExcelName(workBook, "Data Management", true);
Expense_Number = ReadExcelSheet("Expense Number",EnvParams.Opco,"Data Management");
if((Expense_Number=="")||(Expense_Number==null)){
ExcelUtils.setExcelName(workBook, sheetName, true);
Expense_Number = ExcelUtils.getRowDatas("Expense Number",EnvParams.Opco)
} 
if((Expense_Number=="")||(Expense_Number==null)){
 ValidationUtils.verify(true,false,"Expense Number is need to reject expense") 
}

} 
  
  
function goToExpeseMenuItem(){
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.HoverMouse();
ReportUtils.logStep_Screenshot("");
menuBar.DblClick();
if(ImageRepository.ImageSet.TimeExpense.Exists()){
ImageRepository.ImageSet.TimeExpense.Click();
}
else if(ImageRepository.ImageSet.TimeExpense1.Exists()){
ImageRepository.ImageSet.TimeExpense1.Click();
}
else{
ImageRepository.ImageSet.TimeExpense2.Click();
} 
aqUtils.Delay(3000, Indicator.Text);
Sys.Desktop.KeyDown(0x12);
Sys.Desktop.KeyDown(0x20);
Sys.Desktop.KeyUp(0x12);
Sys.Desktop.KeyUp(0x20);
Sys.Desktop.KeyDown(0x58);
Sys.Desktop.KeyUp(0x58);  
aqUtils.Delay(1000, Indicator.Text);
var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
aqUtils.Delay(1000,"Loading Workspace")
var MainBrnch = "";
for(var bi=0;bi<WrkspcCount;bi++){ 
  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
    MainBrnch = Workspc.Child(bi);
    break;
  }
}

var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Time & Expenses").OleValue.toString().trim());
}
}
aqUtils.Delay(10000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Time & Expenses from Time & Expenses Menu");
TextUtils.writeLog("Entering into Time & Expenses from Time & Expenses Menu");
}
  
function gotoTimeExpenses(){ 
    ReportUtils.logStep("INFO","Approve Expenses Second Level is Started");    
    aqUtils.Delay(2000,Indicator.Text); 
    Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Refresh(); 
    var expenses = Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.expensestab;
    expenses.Click();
    ReportUtils.logStep_Screenshot();
    aqUtils.Delay(1000,Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){  } 
    
  var allExpenses = //NameMapping.Sys.Maconomy.ChangeEmployee.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Expense Sheets").OleValue.toString().trim());
   Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", "All Expense Sheets")
   waitForObj(allExpenses)
  allExpenses.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){  } 
  var table = ""
  var sheetno = ""
  var childcount = 0;
  var Add = [];
  var Parent = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  Sys.Process("Maconomy").Refresh()
  for(var i = 0;i<Parent.ChildCount;i++){ 
    if((Parent.Child(i).isVisible()) && (Parent.Child(i).ChildCount == 1)){
    Add[childcount] = Parent.Child(i);
    childcount++;
    }
  }

  Parent = "";
  var pos = 1000;
  for(var i=0;i<Add.length;i++){ 
    if(Add[i].Height<pos){ 
      pos = Add[i].Height;
      Parent = Add[i];
    }
  }


  Log.Message(Parent.FullName)
  table = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table)
  sheetno = Parent.SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  Sys.HighlightObject(sheetno)

   Log.Message(sheetno.FullName) 
   Sys.HighlightObject(sheetno);    
   sheetno.Click();
   sheetno.setText(Expense_Number);
    aqUtils.Delay(1000,Indicator.Text); 
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    var flag=false;  
    for(var v=0;v<table.getItemCount();v++){ 
      if(table.getItem(v).getText_2(1).OleValue.toString().trim()==Expense_Number){ 
        flag=true;
        break;
      }
      else{ 
        table.Keys("[Down]");
      }
     }   
     
     var closefilter = Aliases.Maconomy.Shell2.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl
       waitForObj(closefilter);
       Sys.HighlightObject(closefilter);
       closefilter.Click();
           
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }    
     TextUtils.writeLog("Expense Sheet is available in Maconomy :"+Expense_Number);
    ValidationUtils.verify(flag,true,"Expense Sheet is available in Maconomy"); 
        
        
  var total_curr = 
  Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget;  
  //Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.McGroupWidget.Composite3.McTextWidget
  total_curr.Click();            
  ExcelUtils.setExcelName(workBook,"Data Management", true);
  ExcelUtils.WriteExcelSheet("Expense Total",EnvParams.Opco,"Data Management",total_curr.getText().OleValue.toString().trim());
  Log.Message(total_curr.getText().OleValue.toString().trim());
        
  }
  
  
function print(){
    var print = 
    //Aliases.Maconomy.Group.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.TabFolderPanel.Composite.SWTObject("SingleToolItemControl", "", 5)
   Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite.PTabFolder.TabFolderPanel.SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 5)
   //Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 5)
   
    Sys.HighlightObject(print);
    waitForObj(print)
    ReportUtils.logStep_Screenshot();
    print.Click();    
    
TextUtils.writeLog("Print Expense Sheet is Clicked"); 

  aqUtils.Delay(1000, Indicator.Text);
  WorkspaceUtils.savePDF_And_WriteToExcel("ExpenseMPL","P_ExpenseSheet");
}
  
function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}


function validateExpenseSheet()
{
if(EnvParams.Country.toLowerCase() == "china")
Language = "Chinese (Simplified)";
if(EnvParams.Country.toLowerCase() == "spain")
Language ="Spanish";  
Log.Message(Language)

  ExcelUtils.setExcelName(workBook, "CreateExpense", true);
  var employeeNo= ExcelUtils.getColumnDatas("Employeeno",EnvParams.Opco)
  var expenseCurrency = ExcelUtils.getColumnDatas("currency_1",EnvParams.Opco)

   ExcelUtils.setExcelName(workBook, "Data Management", true);
  var fileName = ExcelUtils.getRowDatas("ExpenseMPL",EnvParams.Opco)
  if((fileName==null)||(fileName=="")){ 
  ValidationUtils.verify(false,true,"ExpenseMPL is needed to validate");
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
 
  var pdflineSplit = docObj.split("\r\n");
  

  var expenseNumber = ExcelUtils.getRowDatas("Expense Number",EnvParams.Opco);
  var expenseAmount = ExcelUtils.getRowDatas("Expense Total",EnvParams.Opco);
  var expenseDescription = ExcelUtils.getRowDatas("Expense Description",EnvParams.Opco);
  var jobNumber  = ExcelUtils.getRowDatas("Job Number",EnvParams.Opco);

  Log.Message(jobNumber)
          
  verifyExpenseNumber(expenseNumber, pdflineSplit);     
  verifyJobNumber(jobNumber, pdflineSplit);       
  verifyEmployeeNumber(employeeNo,pdflineSplit);
  verifyExpenseCurrency(expenseCurrency,pdflineSplit); 
  verifyExpenseDescription(expenseDescription,pdflineSplit);  
  verifyExpenseAmount(expenseAmount, pdflineSplit); 
 
}


function verifyExpenseNumber(expenseNumber,pdflineSplit)
{
    var expenseNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Expense Sheet No").OleValue.toString().trim()))
        {
         if(pdflineSplit[j].includes(expenseNumber))
             {
             Log.Message(expenseNumber+" expenseNumber is matching with Pdf");
             ValidationUtils.verify(true,true,"expenseNumber is is matching with Pdf:"+expenseNumber);
             TextUtils.writeLog("expenseNumber is is matching with Pdf:"+expenseNumber); 
             expenseNoFound = true;
             break;
             }
             }
         if(j==pdflineSplit.length-1 && !expenseNoFound)
          ValidationUtils.verify(false,true,"expenseNumber is not same/found in Expense PDF");
        }  
}

function verifyEmployeeNumber(employeeNo,pdflineSplit)
{
  var employeeNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Employee No").OleValue.toString().trim()))
        {
         if(pdflineSplit[j].includes(employeeNo))
             {
             Log.Message(employeeNo+" EmployeeNumber is matching with Pdf");
             ValidationUtils.verify(true,true," EmployeeNumber is matching with Pdf:"+employeeNo);
             TextUtils.writeLog(" EmployeeNumber is matching with Pdf:"+employeeNo);
             employeeNoFound = true;
             break;
             }
             }
         if(j==pdflineSplit.length-1 && !employeeNoFound)
          ValidationUtils.verify(false,true,"EmployeeNumber is not same/found in Expense PDF");
        }  
}

function verifyExpenseAmount(expenseAmount,pdflineSplit)
{
    var expenseAmountFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(expenseAmount))
             {
             Log.Message(expenseAmount+"expenseAmount is matching with Pdf");
             ValidationUtils.verify(true,true," expenseAmount is matching with Pdf:"+expenseAmount);
             TextUtils.writeLog(" expenseAmount is matching with Pdf:"+expenseAmount);
             expenseAmountFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !expenseAmountFound)
          ValidationUtils.verify(false,true,"expenseAmount is not same/found in Expense PDF");
        }  
}
 

function verifyExpenseCurrency(expenseCurrency,pdflineSplit)
{
    var expenseCurrencytFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(expenseCurrency))
             {
             Log.Message(expenseCurrency+"expenseCurrency is matching with Pdf");
             ValidationUtils.verify(true,true," expenseCurrency is matching with Pdf:"+expenseCurrency);
             TextUtils.writeLog(" expenseCurrency is matching with Pdf:"+expenseCurrency);
             expenseCurrencytFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !expenseCurrencytFound)
          ValidationUtils.verify(false,true,"expenseCurrency is not same/found in Expense PDF");
        }  
}

function verifyExpenseDescription(expenseDescription,pdflineSplit)
{
    var expenseDescriptionFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(expenseDescription))
             {
             Log.Message(expenseDescription+" expenseDescription is matching with Pdf");
             ValidationUtils.verify(true,true," expenseDescription is matching with Pdf:"+expenseDescription);
             TextUtils.writeLog(" expenseDescription is matching with Pdf:"+expenseDescription);
             expenseDescriptionFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !expenseDescriptionFound)
          ValidationUtils.verify(false,true,"expenseDescription is not same/found in Expense PDF");
        }  
}
 
function verifyJobNumber(jobNumber,pdflineSplit)
{
    var jobNoFound = false;
  for (var j=0; j<pdflineSplit.length; j++)
  {
         if(pdflineSplit[j].includes(jobNumber))
             {
             Log.Message(jobNumber+" jobNumber is matching with Pdf");
             ValidationUtils.verify(true,true," jobNumber is matching with Pdf:"+jobNumber);
             TextUtils.writeLog("jobNumber is matching with Pdf:"+jobNumber);
             jobNoFound = true;
             break;
             }
         if(j==pdflineSplit.length-1 && !jobNoFound)
          ValidationUtils.verify(false,true,"jobNumber is not same/found in Expense PDF");
        }  
}