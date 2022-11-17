//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart


//-------Data Sheet path----------
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;
var sheetName = "Payment Listing";
var Language = "";
Indicator.Show();


//-------Global Variables----------
var VendorNo, Duedate, Paymentagent, Paymentmode, layoutTypes = "";



//-------Main Function----------
//   Triggering to do payment for "Manual mode" vendor invoice
//   User Role Required: SSC - Senior AP
//   Payment Selection need to be created for vendor invoice before triggering payment listing

function paymentListing(){ 

	//Language to execute the test case
	TextUtils.writeLog("Void a Payment Started"); 
	Indicator.PushText("waiting for window to open");
	Language = "";
	Language = EnvParams.LanChange(EnvParams.Language);
	WorkspaceUtils.Language = Language;

	// Validating and restart application in required user role
	var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
	menuBar.Click();
	ExcelUtils.setExcelName(workBook, "SSC Users", true);
	var Project_manager = ExcelUtils.getRowDatas("SSC - Senior AP","Username")
	if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
	WorkspaceUtils.closeMaconomy();
	Restart.login(Project_manager);
	  
	}


	VendorNo, Duedate, Paymentagent, Paymentmode, layoutTypes = "";

	try{
	getDetails();
	goToBankingTransaction();  
  paymentListing_Tab(); 
	entring_selction_criteria(); 
	printLayout(); 
	}
	  catch(err){
	    Log.Message(err);
	  }

	}

//	    Getting required data from datasheet
 function getDetails(){

	//  Vendor No
	ExcelUtils.setExcelName(workBook, "Data Management", true);
	VendorNo = ReadExcelSheet("Vendor Number",EnvParams.Opco,"Data Management");
	Log.Message(VendorNo)
	if((VendorNo=="")||(VendorNo==null)){
	ExcelUtils.setExcelName(workBook, sheetName, true);
	VendorNo = ExcelUtils.getRowDatas("Vendor Number",EnvParams.Opco)
	}
	if((VendorNo==null)||(VendorNo=="")){ 
	ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Listing");
	}

	//  Due Date
	ExcelUtils.setExcelName(workBook, "Data Management", true);
	Duedate = ReadExcelSheet("Vendor Invoice Due Date",EnvParams.Opco,"Data Management");
	if((Duedate=="")||(Duedate==null)){
	ExcelUtils.setExcelName(workBook, sheetName, true);
	Duedate = ExcelUtils.getRowDatas("DueDate",EnvParams.Opco)
	}
	Log.Message(Duedate)
	if((Duedate==null)||(Duedate=="")){ 
	ValidationUtils.verify(false,true,"Due Date Number is Needed to Create a Payment Listing");
	}

	//  Payment Agent
	ExcelUtils.setExcelName(workBook, sheetName, true);
	Paymentagent = ExcelUtils.getRowDatas("Payment_Agent",EnvParams.Opco)
	Log.Message(Paymentagent)
	if((Paymentagent==null)||(Paymentagent=="")){ 
	ValidationUtils.verify(false,true,"Payment Agent is Needed to Create a Payment Listing");
	}

	//  Payment Mode
	ExcelUtils.setExcelName(workBook, "Data Management", true);
	Paymentmode = ReadExcelSheet("Vendor Invoice Payment Mode",EnvParams.Opco,"Data Management");
	if((Paymentmode=="")||(Paymentmode==null)){
	ExcelUtils.setExcelName(workBook, sheetName, true);
	Paymentmode = ExcelUtils.getRowDatas("Paymode_Mode",EnvParams.Opco) 
	}
	Log.Message(Paymentmode)
	if((Paymentmode==null)||(Paymentmode=="")){ 
	ValidationUtils.verify(false,true,"Paymode Mode Number is Needed to Create a Payment Listing");
	}
//Layout
ExcelUtils.setExcelName(workBook, sheetName, true);
layoutTypes = ExcelUtils.getRowDatas("Layout",EnvParams.Opco)
Log.Message(layoutTypes)
if((layoutTypes==null)||(layoutTypes=="")){ 
ValidationUtils.verify(false,true,"Layout is Needed to Create a Payment Listing");
}
	}

	// Navigating Banking > Bank transactions
	function goToBankingTransaction(){

	var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
	menuBar.HoverMouse();
	ReportUtils.logStep_Screenshot("");
	menuBar.DblClick();
	if(ImageRepository.ImageSet.Banking.Exists()){
	 ImageRepository.ImageSet.Banking.Click();// GL
	}
	else if(ImageRepository.ImageSet.Banking1.Exists()){
	ImageRepository.ImageSet.Banking1.Click();
	}
	else{
	ImageRepository.ImageSet.Banking2.Click();
	}

	var WrkspcCount = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
	var Workspc = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
	var MainBrnch = "";
	for(var bi=0;bi<WrkspcCount;bi++){ 
	  if((Workspc.Child(bi).isVisible())&&(Workspc.Child(bi).Child(0).Name.indexOf("Composite")!=-1)&&(Workspc.Child(bi).Child(0).isVisible())){ 
	    MainBrnch = Workspc.Child(bi);
	    break;
	  }
	}


	var childCC= MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
	var Banking;

	for(var i=1;i<=childCC;i++){ 
	Banking = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
	if(Banking.isVisible()){ 
	Banking = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
	Banking.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions").OleValue.toString().trim());
	ReportUtils.logStep_Screenshot();
	Banking.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Bank Transactions").OleValue.toString().trim());
	}
	}

	ReportUtils.logStep("INFO", "Moved to Banking Transactions from job Menu");
	TextUtils.writeLog("Entering into Banking Transactions from Jobs Menu");
	}

	// Navigating Bank transactions > Bank Payments > Payment > Payment Listing
	function paymentListing_Tab() {

		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}

		var payment = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Payment;
		Sys.HighlightObject(payment);
		payment.Click();

		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");

		var payemntListing = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.TabControl;
		Sys.HighlightObject(payemntListing);
		payemntListing.Click();

		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");

	}

//	    Tick the Approve Selection and Show Entries fields. 
//	    Enter desired Selection Criteria.
//	    In Print Control island, select Layout 'WPP PaymentList'
//	    Then click 'Print'
	function entring_selction_criteria(){ 
	  
	var approveSelection = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.McPlainCheckboxView.Button;
	Sys.HighlightObject(approveSelection);

	    if(!approveSelection.getSelection()){
	      approveSelection.Click();
	      ValidationUtils.verify(true,true,"Approve Selection is Ticked");
	      checkmark = true;
	    }else{ 
        ValidationUtils.verify(true,true,"Approve Selection is Ticked");
      }
	    
	var showEntries = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite2.McPlainCheckboxView.Button;
	Sys.HighlightObject(showEntries);

	    if(!showEntries.getSelection()){
	      showEntries.Click();
	      ValidationUtils.verify(true,true,"Show Entries is Ticked");
	      checkmark = true;
	    }else{ 
        ValidationUtils.verify(true,true,"Show Entries is Ticked");
      }
	   
	var preferedDate = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite3.McDatePickerWidget;
	Sys.HighlightObject(preferedDate);
	  if((Duedate!="")&&(Duedate!=null)){
	       aqUtils.Delay(1000, Indicator.Text);
	       preferedDate.setText(Duedate);
	          ValidationUtils.verify(true,true,"Due Date is selected in Maconomy"); 
	        }
	    else{ 
	      ValidationUtils.verify(false,true,"Payment Date is Needed  for Payment Listing");
	    }

  var payemntAgent = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite.McValuePickerWidget;
  if(Paymentagent!=""){
  payemntAgent.Click();
  WorkspaceUtils.SearchByValue(payemntAgent,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Agent").OleValue.toString().trim(),Paymentagent,"Payment Agent")
  ValidationUtils.verify(true,true,"Payment Agent is selected in Maconomy"); 
  }
  else{ 
    ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment Listing");
  }
  
  var paymentMode = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite2.McValuePickerWidget;
  if(Paymentmode!=""){
  paymentMode.Click();
  WorkspaceUtils.SearchByValue(paymentMode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Payment Mode").OleValue.toString().trim(),Paymentmode,"Payment Mode")
  ValidationUtils.verify(true,true,"Payment Mode is selected in Maconomy"); 
  }
  else{ 
    ValidationUtils.verify(false,true,"Payment Agent is Needed to Create Payment Listing");
  }
  
  var vendorFrom = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite3.McValuePickerWidget;
  Sys.HighlightObject(vendorFrom);
  if(VendorNo!=""){
  vendorFrom.Click();
  WorkspaceUtils.VPWSearchByValue(vendorFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
  ValidationUtils.verify(true,true,"Vendor No From is selected in Maconomy"); 
 }
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Listing");
  }
  
  var vendorTo = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite3.McValuePickerWidget2;
  Sys.HighlightObject(vendorTo);
  if(VendorNo!=""){
  vendorTo.Click();
  WorkspaceUtils.VPWSearchByValue(vendorTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Vendor").OleValue.toString().trim(),VendorNo,"Vendor Number");
 ValidationUtils.verify(true,true,"Vendor No To is selected in Maconomy"); 
 }
 
 else{ 
    ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Payment Listing");
  }
  
  var companyFrom = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite4.McValuePickerWidget;
  Sys.HighlightObject(companyFrom)
  companyFrom.Click();
  WorkspaceUtils.SearchByValue(companyFrom,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
  ValidationUtils.verify(true,true,"Company No From is selected in Maconomy"); 
  aqUtils.Delay(1000, Indicator.Text);
  
  var companyTo = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.Composite4.McValuePickerWidget2;
	Sys.HighlightObject(companyTo)
  companyTo.Click();
  WorkspaceUtils.SearchByValue(companyTo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Company").OleValue.toString().trim(),EnvParams.Opco,"Company Number");
  ValidationUtils.verify(true,true,"Company No To is selected in Maconomy"); 
  aqUtils.Delay(1000, Indicator.Text);   
	    
	}
  
  // Select Layout and Print
  function printLayout(){ 
    
  if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
    
    var layout = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite3.McGroupWidget.Composite.McPopupPickerWidget;
  Log.Message(layoutTypes)
  layout.Keys(layoutTypes);
  
  aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
    
  //Save Changes
  var save = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  waitForObj(save)
  Sys.HighlightObject(save)
  save.Click();
  ReportUtils.logStep_Screenshot();
		if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
  
    ValidationUtils.verify(true,true,"Selection Criteria is entered in Maconomy"); 
    
  //Print payment
  var print = Aliases.Maconomy.Payment_Listing.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
  waitForObj(print)
  print.Click();
  ValidationUtils.verify(true, true, "Print Icon is clicked")
  if (ImageRepository.ImageSet.Tab_Icon.Exists()) {
		}
		aqUtils.Delay(4000, "Validing Maconomy Screen is loaded completly with full information");
  //Save PDF in Local Directory
  savePDF_And_WriteToExcel("Print Listing");
  
  }



