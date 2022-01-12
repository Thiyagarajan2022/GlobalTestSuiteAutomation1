//USEUNIT EnvParams
//USEUNIT ExcelUtils
//USEUNIT ReportUtils
//USEUNIT TestRunner
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT Restart

Indicator.Show();
var excelName = EnvParams.path;
var workBook = Project.Path+excelName;

var Project_manager="";
var level =0;
var STIME = "";
var Invoice_Editing_Number,Job_Number = "";
var Language = "";
var URL,Hitpoint_User_Name,Hitpoint_Password = "";


//Main Function
function Hitpoint_For_Client_Invoice(){ 
TextUtils.writeLog("Create Invoice Intergration with Hitpoint Started"); 
Indicator.PushText("waiting for window to open");
aqUtils.Delay(1000, Indicator.Text);
Language = EnvParams.LanChange(EnvParams.Language);
WorkspaceUtils.Language = Language;


aqUtils.Delay(1000, Indicator.Text);
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.Click();
ExcelUtils.setExcelName(workBook, "SSC Users", true);
var Project_manager = ExcelUtils.getRowDatas("SSC - Biller","Username")
if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption.toString().trim().indexOf(Project_manager)==-1){ 
WorkspaceUtils.closeMaconomy();
Restart.login(Project_manager);
  
}
URL,Hitpoint_User_Name,Hitpoint_Password = "";
STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Hitpoint Intergration started::"+STIME);
TextUtils.writeLog("Execution Start Time :"+STIME); 

getDetails();
goto_Hitpoint_Billing();
var Maconomy_Status = Import_to_Hitpoint_Successful();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.Click();
WorkspaceUtils.closeAllWorkspaces();
if(Maconomy_Status){ 
  Hitpoint_integration();

}

}


function getDetails(){ 
  
ExcelUtils.setExcelName(workBook, "Data Management", true);
Invoice_Editing_Number = ReadExcelSheet("Invoice Editing Number",EnvParams.Opco,"Data Management");
if((Invoice_Editing_Number=="")||(Invoice_Editing_Number==null)){
  ValidationUtils.verify(false,true,"Invoice_Editing_Number is Missing in Data Mangementto integrate with Hitpoint");
}
Log.Message(Invoice_Editing_Number)

Job_Number = Invoice_Editing_Number.substring(Invoice_Editing_Number.indexOf("#")+1,Invoice_Editing_Number.lastIndexOf("#"))
Log.Message(Job_Number)


ExcelUtils.setExcelName(workBook, "HitPoint", true);
URL,Hitpoint_User_Name,Hitpoint_Password = "";
 URL = ExcelUtils.getRowDatas("HitPoint URL","Value")
if((URL==null)||(URL=="")){ 
ValidationUtils.verify(false,true,"Hitpoint URL is Needed to Create  Client Invoice");
}

 Hitpoint_User_Name = ExcelUtils.getRowDatas("User Name","Value")
if((Hitpoint_User_Name==null)||(Hitpoint_User_Name=="")){ 
ValidationUtils.verify(false,true,"Hitpoint_User_Name is Needed to Create  Client Invoice");
}

 Hitpoint_Password = ExcelUtils.getRowDatas("Password","Value")
if((Hitpoint_Password==null)||(Hitpoint_Password=="")){ 
ValidationUtils.verify(false,true,"Hitpoint_Password is Needed to Create Client Invoice");
}


}




function goto_Hitpoint_Billing(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
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
  var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");

Log.Message(Language)
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Hitpoint Billing").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Hitpoint Billing from Jobs Menu");
TextUtils.writeLog("Entering into Hitpoint Billing from Jobs Menu");


}


function Import_to_Hitpoint_Successful(){ 
aqUtils.Delay(4000,"Checking Hitpoint Billing Status");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Checking Hitpoint Billing Status");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
var companyNo =   Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(companyNo);
companyNo.Click();
companyNo.setText(EnvParams.Opco);

aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");

var Editing_Number =   Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Editing_Number);
Editing_Number.Click();
Editing_Number.setText(Invoice_Editing_Number);

aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");


aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");

aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");

var Job_No =   Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(Job_No);
Job_No.Click();
Job_No.setText(Job_Number);

var table =  Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if((table.getItem(v).getText_2(1).OleValue.toString().trim()==Invoice_Editing_Number) && (table.getItem(v).getText_2(4).OleValue.toString().trim()==Job_Number)){ 
      
      if(table.getItem(v).getText_2(7).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Import to Hitpoint Successful").OleValue.toString().trim()){
      ValidationUtils.verify(true,true,"Status changed in Maconomy as Import to Hitpoint Successful")  
      return true;
      }
      else if(table.getItem(v).getText_2(7).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice posted and Completed").OleValue.toString().trim()){
      ValidationUtils.verify(true,true,"Status changed in Maconomy as Invoice posted and Completed");
      var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
      menuBar.Click();
      WorkspaceUtils.closeAllWorkspaces();
      validate_Maconomy_Status();
      return false;
      }
      else{
      ValidationUtils.verify(false,true,"Status is NOT CHANGED in Maconomy as Import to Hitpoint Successful");
        return false;
      }
      flag=true;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

}


function Hitpoint_integration(){ 
  
Browsers.Item(btChrome).Run(URL);

var page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login");
var User_Name = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(1).Textbox("userName");
User_Name.SetText(Hitpoint_User_Name);
var Password = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(2).PasswordBox("password")
Password.SetText(Hitpoint_Password);

var CaptchaVar = BuiltIn.InputBox("Captcha", "Please enter a CAPTCHA showing in the browser", "");
var Captcha = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(3).Textbox("verifyCodeId");
Captcha.SetText(CaptchaVar);

var Login = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/login").Panel(0).Panel(1).Panel(4).Button("btnSysLogin");
Login.Click();

page.Wait();


 // Finds the Link object on the page
  var link = page.NativeWebObject.Find("href", "http://101.231.221.6:8090/wppoutput/base/sys/main#1", "A");

  // If the link is found
  if (link.Exists)
  {
    // Clicks the link
    link.Click();
    // Waits until the target page is loaded
    page.Wait();
  }
  


var Switch_To_English = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(0).Panel(0).Panel(0).Panel(1).Link("tms_sys_select_Language_a_en_US");
Switch_To_English.Click();

var entity = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(0).Panel(0).Panel(0).Panel(1).Link("tms_sys_select_stxx");         
entity.Click();

var Company_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel("dynamic_criteria_form_sys_selectst").Panel("querydivid").Panel(0).Form("criteriaQueryForm_sys_selectst").Table(0).Cell(0, 5).Textbox("easyui_textbox_input4");
Company_Number.Click();

Company_Number.SetText("1221");

var search = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel("dynamic_criteria_form_sys_selectst").Panel(0).Link("btn_criteriaForm_select_sys_selectst").TextNode(1);
search.Click();
aqUtils.Delay(5000, "Waiting to load");

var table = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1);
var TableSize = table.ChildCount;
Log.Message(TableSize)

var OpCo_Status = false;
for (var i=1;i<TableSize;i++){ 
Log.Message(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(i).Table(0).Cell(0, 4).Panel(0).contentText)
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(i).Table(0).Cell(0, 4).Panel(0).contentText=="1221"){ 
  var Check_Box = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1", 0).Panel(6).Panel("main_sys_changeCurrentSTXX_dialog").Panel("dynamic_criteria_content_area_sys_selectst").Panel("dynamic_criteria_table_area_sys_selectst").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(i).Table(0).Cell(0, 0).Panel(0).Checkbox("ck");
  Sys.HighlightObject(Check_Box);
  Check_Box.Click();
  aqUtils.Delay(3000, "Waiting to load");
  var OKay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel(1).Link("tms_sys_stxx_dlg_confirm").TextNode(1);
  Sys.HighlightObject(OKay);
  OKay.Click();
  OpCo_Status = true;
  var Pop_Up_Message = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel(1).Panel(1);
  Log.Message(Pop_Up_Message.contentText)
  var Pop_Up_Okay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(6).Panel(2).Link(0).TextNode(0);
  Pop_Up_Okay.Click();

break;
}

}

if(!OpCo_Status){ 
  Log.Error("OpCo is not available in entity");
}


aqUtils.Delay(5000, "Waiting to load");
var Select_Module = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(0).Panel(0).Panel(0).Panel(0).Table(0).Textbox("easyui_textbox_input1");
Select_Module.Click();
var Output_Management = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(5).Panel(0).Panel("easyui_combobox_i1_1");
Output_Management.Click();
aqUtils.Delay(5000, "Selecting Opertation Center");

var Invoice_management = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_19").TextNode(0);
Invoice_management.Click();
var Operation_center = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_20").TextNode(0);
Operation_center.Click();
aqUtils.Delay(5000, "Selecting Opertation Center");
var Document_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_form_200201").Panel("querydivid").Panel(0).Form("criteriaQueryForm_200201").Table(0).Cell(0, 1).Textbox("easyui_textbox_input5");
Document_Number.Click();
Document_Number.SetText(Invoice_Editing_Number);

var Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_form_200201").Panel(1).Link("btn_criteriaForm_select_200201").TextNode(0);
Query.Click();
aqUtils.Delay(8000, "Selecting Document Number");

var total_page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 5).TextNode(0);
Log.Message(total_page.textContent);
total_page = total_page.textContent;
total_page = total_page.substring(total_page.indexOf("of ")+3);
Log.Message(total_page);

var table = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0);
Sys.HighlightObject(table);

var Invoice_Status = false;
var Document_Nums = [];
var j = 0;
var Document_Satus = false;

for (var k=0;k<total_page;k++){ 
var table_Grid = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).RowCount;
for (var i=0;i<table_Grid;i++){  
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 5).contentText=="Save in Store"){ 
  var temp = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 1).Panel(0);
  Document_Nums[j] = temp.contentText;
  Log.Message(temp.contentText)
  Invoice_Status = true;
  j++;
  }
  else if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 5).contentText=="Invoiced"){ 
    Document_Satus = true;
  }
  }

  var Next_Page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 6).Link(0).TextNode(1);
  Sys.HighlightObject(Next_Page);
  Next_Page.Click();
aqUtils.Delay(5000, "Waiting to load");
  }
  
  
  if(!Document_Satus){
  Log.Message(" ");
  for (var k=0;k<Document_Nums.length;k++){ 
    Log.Message(Document_Nums[k]);
  }
  
var Switch_To_Chinese = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(0).Panel(0).Panel(0).Panel(1).Link("tms_sys_select_Language_a_zh_CN");
Switch_To_Chinese.Click();

aqUtils.Delay(5000, "Waiting to load");
var Select_Module = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(0).Panel(0).Panel(0).Panel(0).Table(0).Textbox("easyui_textbox_input1");
Select_Module.Click();
var Output_Management = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(5).Panel(0).Panel("easyui_combobox_i1_1");
Output_Management.Click();
aqUtils.Delay(5000, "Selecting Opertation Center");

var Invoice_management = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_19").TextNode(0);
Invoice_management.Click();
var Operation_center = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_20").TextNode(0);
Operation_center.Click();
aqUtils.Delay(5000, "Selecting Opertation Center");
var Document_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_form_200201").Panel("querydivid").Panel(0).Form("criteriaQueryForm_200201").Table(0).Cell(0, 1).Textbox("easyui_textbox_input2")
Document_Number.Click();
Document_Number.SetText(Invoice_Editing_Number);

var Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_form_200201").Panel(1).Link("btn_criteriaForm_select_200201").TextNode(0);
Query.Click();
aqUtils.Delay(8000, "Selecting Document Number");
Log.Message("total_page :"+total_page);

var Review_Status = false;
for (var j=0;j<total_page;j++){ 
var table_Grid = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).RowCount;
Log.Message("table_Grid :"+table_Grid);
for (var i=0;i<table_Grid;i++){  
  
for (var k=0;k<Document_Nums.length;k++){  
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 1).Panel(0).contentText==Document_Nums[k]){ 

   var Check_Box = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 0).Panel(0).Checkbox("ck");
   Sys.HighlightObject(Check_Box);
   Check_Box.Click();
   Review_Status = true;
   aqUtils.Delay(4000, "Selecting Document Number");
   break;
  }
  }
  }
  var Next_Page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 6).Link(0).TextNode(1);
  Sys.HighlightObject(Next_Page);
  Next_Page.Click();
aqUtils.Delay(5000, "Waiting to load");
  }
  aqUtils.Delay(5000, "Selecting Document Number");
  
  if(Review_Status){
  var Review = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_toolbar_200201").Panel("buttondivid").Link("dynamic_criteria_button_20020102").TextNode(1);
  Sys.HighlightObject(Review);
  Review.Click();
  
  aqUtils.Delay(3000, "Waiting to load");
  
  var Notification_Msg = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(17).Panel(1).Panel(1);
  Log.Message(Notification_Msg.contentText)
  var Okay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(17).Panel(2).Link(0).TextNode(0);
  Okay.Click();
  }
  aqUtils.Delay(5000, "Waiting to load");
  
  var Change_To_English = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(0).Panel(0).Panel(0).Panel(1).Link("tms_sys_select_Language_a_en_US");
  Change_To_English.Click();
  aqUtils.Delay(5000, "Waiting to load");

var Select_Module = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(0).Panel(0).Panel(0).Panel(0).Table(0).Textbox("easyui_textbox_input1");
Select_Module.Click();
var Output_Management = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(5).Panel(0).Panel("easyui_combobox_i1_1");
Output_Management.Click();
aqUtils.Delay(5000, "Selecting Opertation Center");

var Invoice_management = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_19").TextNode(0);
Invoice_management.Click();
var Operation_center = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_20").TextNode(0);
Operation_center.Click();
aqUtils.Delay(5000, "Selecting Opertation Center");
var Document_Number = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_form_200201").Panel("querydivid").Panel(0).Form("criteriaQueryForm_200201").Table(0).Cell(0, 1).Textbox("easyui_textbox_input2");
Document_Number.Click();
Document_Number.SetText(Invoice_Editing_Number);

var Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel("dynamic_criteria_form_200201").Panel(1).Link("btn_criteriaForm_select_200201");
Query.Click();
aqUtils.Delay(5000, "Selecting Document Number");

var total_page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 5);
Log.Message(total_page.textContent);
total_page = total_page.textContent;
total_page = total_page.substring(total_page.indexOf("of ")+3);
Log.Message(total_page);

var table = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0);
Sys.HighlightObject(table);

var Invoice_Status = false;



for (var k=0;k<total_page;k++){ 
var table_Grid = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).RowCount;
for (var i=0;i<table_Grid;i++){  
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 5).Panel(0).contentText!="Reviewed"){ 
ValidationUtils.verify(false,true,"Document Number is not Reviewed");
  }
  }
  
  var Next_Page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 6).Link(0).TextNode(1);
  Sys.HighlightObject(Next_Page);
  Next_Page.Click();
aqUtils.Delay(5000, "Waiting to load");
  }
  

  
  var Invoice_tab = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(1).Panel(0).Panel(0).Panel("easyui_tree_22").TextNode(0);
  Invoice_tab.Click();
  aqUtils.Delay(5000, "Waiting to load");
  
  var Document_Num = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel("dynamic_criteria_form_200204").Panel("querydivid").Panel(0).Form("criteriaQueryForm_200204").Table(0).Cell(0, 1).Textbox("easyui_textbox_input25");
  Sys.HighlightObject(Document_Num);
  Document_Num.Click();
  Document_Num.SetText(Invoice_Editing_Number);
  
  var Query = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel("dynamic_criteria_form_200204").Panel(1).Link("btn_criteriaForm_select_200204").TextNode(0);
  Sys.HighlightObject(Query);
  Query.Click();
  aqUtils.Delay(5000, "Waiting to load");
  
  var All_Document_CheckBox = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(0).Panel(0).Table(0).Cell(0, 0).Panel(0).Checkbox(0);
  Sys.HighlightObject(All_Document_CheckBox);
  All_Document_CheckBox.Click();
  aqUtils.Delay(3000, "Waiting to load");
  
  var Opening = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel("dynamic_criteria_toolbar_200204").Panel("buttondivid").Link("dynamic_criteria_button_20020401").TextNode(0);
  Sys.HighlightObject(Opening);
  Opening.Click();
  aqUtils.Delay(5000, "Waiting to load");
  
  var Heading = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(0).Panel(0);
  if(Heading.innerText=="Tips"){
  Notification_Msg = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(1).Panel(1);
  Log.Message(Notification_Msg.contentText)
  var OKay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(2).Link(0)
  OKay.Click();
  aqUtils.Delay(8000, "Waiting to load");  
  }
  
//  Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(0).Panel(0)
  var Save = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(1).Link("openinv_make_dlg_submit").TextNode(1);
  Save.Click();
  aqUtils.Delay(8000, "Waiting to load"); 
  var Heading = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(0).Panel(0);
  if(Heading.innerText=="Tips"){
  Notification_Msg = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(1).Panel(1);
  Log.Message(Notification_Msg.contentText)
  Okay = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(14).Panel(2).Link(0).TextNode(0);
  OKay.Click();
  aqUtils.Delay(8000, "Waiting to load");  
  }
  
  
  total_page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 5).TextNode(0);
Log.Message(total_page.textContent);
total_page = total_page.textContent;
total_page = total_page.substring(total_page.indexOf("of ")+3);
Log.Message(total_page);
var Review_Status = false;
for (var j=0;j<total_page;j++){ 
var table_Grid = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).RowCount;
Log.Message("table_Grid :"+table_Grid);
for (var i=0;i<table_Grid;i++){  
  
for (var k=0;k<Document_Nums.length;k++){  
if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 1).Panel(0).contentText==Document_Nums[k]){ 

if(Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200204").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200204").Panel("dynamic_criteria_table_area_200204").Panel(0).Panel(0).Panel(0).Panel(0).Panel(1).Panel(1).Table(0).Cell(i, 4).Panel(0).contentText!="Invoiced")
  ValidationUtils.verify(false,true,"Document Number is not Invoiced");
  }
  }
  }
  var Next_Page = Sys.Browser("chrome").Page("http://101.231.221.6:8090/wppoutput/base/sys/main#1&200201").Panel(2).Panel("systemMainBodyPageContextArea").Panel("dynamic_criteria_content_area_200201").Panel("dynamic_criteria_table_area_200201").Panel(0).Panel(0).Panel(0).Panel(1).Table(0).Cell(0, 6).Link(0).TextNode(1);
  Sys.HighlightObject(Next_Page);
  Next_Page.Click();
aqUtils.Delay(5000, "Waiting to load");
  }
  aqUtils.Delay(5000, "Selecting Document Number");
  
  
  }
else{
validate_Maconomy_Status();
}


}




function Invoice_posted_and_Completed_Successful(){ 
aqUtils.Delay(4000,"Checking Hitpoint Billing Status");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
aqUtils.Delay(4000,"Checking Hitpoint Billing Status");
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }

var companyNo =   Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(companyNo);
companyNo.Click();
companyNo.setText(EnvParams.Opco);

aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");

var Editing_Number =   Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
Sys.HighlightObject(Editing_Number);
Editing_Number.Click();
Editing_Number.setText(Invoice_Editing_Number);

aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");


aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");

aqUtils.Delay(2000,"Checking Hitpoint Billing Status");
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
aqUtils.Delay(2000,"Checking Hitpoint Billing Status");

var Job_No =   Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
Sys.HighlightObject(Job_No);
Job_No.Click();
Job_No.setText(Job_Number);

var table =  Aliases.Maconomy.Hitpoint.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
Sys.HighlightObject(table);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if((table.getItem(v).getText_2(1).OleValue.toString().trim()==Invoice_Editing_Number) && (table.getItem(v).getText_2(4).OleValue.toString().trim()==Job_Number)){ 
      if(table.getItem(v).getText_2(7).OleValue.toString().trim()==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Invoice posted and Completed").OleValue.toString().trim()){
      ValidationUtils.verify(true,true,"Status changed in Maconomy as Invoice posted and Completed");
      return true;
      }
      else{
      ValidationUtils.verify(false,true,"Status is NOT CHANGED in Maconomy as Invoice posted and Completed");
        return false;
      }
      flag=true;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

}



function goto_Jobs(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
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
  var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = MainBrnch.SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");

Log.Message(Language)
Client_Managt.ClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
ReportUtils.logStep_Screenshot();
Client_Managt.DblClickItem("|"+JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Jobs").OleValue.toString().trim());
}

} 

//aqUtils.Delay(5000, Indicator.Text);
ReportUtils.logStep("INFO", "Moved to Jobs from Jobs Menu");
TextUtils.writeLog("Entering into Jobs from Jobs Menu");
}


function Print_Invoice(){  
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  aqUtils.Delay(3000,"Maconomy Loading data")
     TextUtils.writeLog("Customer Payment for Single Invoice is started");
      ReportUtils.logStep("INFO", "Customer Payment for Single Invoice is started::"+STIME);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var allJobs = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
  WorkspaceUtils.waitForObj(allJobs);
  allJobs.Click();

  var table = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid;
  var firstcell = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McValuePickerWidget;
  var closeFilter = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
  WorkspaceUtils.waitForObj(firstcell);
  firstcell.forceFocus();
  firstcell.setVisible(true);
  firstcell.ClickM();
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyDown(0x09);
  aqUtils.Delay(1000, Indicator.Text);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  
  var job = Aliases.Maconomy.WorkingEstimate.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McTableWidget.McGrid.McTextWidget;
  job.Click();
  job.setText(Job_Number);
  WorkspaceUtils.waitForObj(job);
  WorkspaceUtils.waitForObj(table);
//  aqUtils.Delay(7000, Indicator.Text);
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(2).OleValue.toString().trim()==Job_Number){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }
  
//  if(flag){
  ReportUtils.logStep("INFO", "Job is listed in table to Print Invoice");
  ReportUtils.logStep_Screenshot("");
  TextUtils.writeLog("Job("+Job_Number+") is available in maconommy to Print Invoice"); 
  closeFilter.Click();
  
if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
}

                        
       var invoice = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.relocation;
       Sys.HighlightObject(invoice);       
       ReportUtils.logStep_Screenshot(""); 
       invoice.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
       
       var invoicehistory = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.TabFolderPanel.history;
       Sys.HighlightObject(invoicehistory);
       invoicehistory.Click();
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
  }
       var table = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite9.Composite.PTabFolder.Composite3.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McTableWidget.invoicetable_1;
       var Invoice_Number = table.getItem(0).getText_2(0).OleValue.toString().trim();
       
       var Print_Invoice_Copy = Aliases.Maconomy.Screen3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
       Sys.HighlightObject(Print_Invoice_Copy);
       Print_Invoice_Copy.Click();
       
      TextUtils.writeLog("Print Client Invoice is Clicked and saved"); 
    aqUtils.Delay(9000, Indicator.Text);
var SaveTitle = "";
var sFolder = "";
var pdf = Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVTopBarView", 4);
    if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*Invoice"+"*", 1).WndCaption.indexOf("Invoice")!=-1){
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x12); 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x41);
    
    if(ImageRepository.PDF.ChooseFolder.Exists())
    ImageRepository.PDF.ChooseFolder.Click();
    else{ 
      var window = Sys.Process("AcroRd32", 2).Window("AVL_AVDialog", "Save As", 1).Window("AVL_AVView", "AVAiCDialogView", 1);
      WorkspaceUtils.waitForObj(window);
      Sys.Desktop.KeyDown(0x12); //Alt
      Sys.Desktop.KeyDown(0x73); //F4
      Sys.Desktop.KeyUp(0x12); //Alt
      Sys.Desktop.KeyUp(0x73); //F4
    aqUtils.Delay(2000, Indicator.Text);
    Sys.HighlightObject(pdf)
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x41); //A 
    Sys.Desktop.KeyUp(0x12); 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x41);
    }
    var save = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).UIAObject("Explorer_Pane").Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1);
    aqUtils.Delay(2000, Indicator.Text);
    SaveTitle = save.wText;
    
sFolder = Project.Path+"MPLReports\\"+EnvParams.TestingType+"\\"+EnvParams.Country+"\\"+EnvParams.Opco+"\\";
if (! aqFileSystem.Exists(sFolder)){
if (aqFileSystem.CreateFolder(sFolder) == 0){ 
    
}
else{
Log.Error("Could not create the folder " + sFolder);
}
}
save.Keys(sFolder+SaveTitle+".pdf");
//var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1).Window("Button", "&Save", 1);
//saveAs.Click();
var saveAs = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
var p = Sys.Process("AcroRd32").Window("#32770", "Save As", 1);
Sys.HighlightObject(p);
var saveAs = p.FindChild("WndCaption", "&Save", 2000);
if (saveAs.Exists)
{ 
saveAs.Click();
}
aqUtils.Delay(2000, Indicator.Text);
if(ImageRepository.ImageSet.SaveAs.Exists()){
var conSaveAs = Sys.Process("AcroRd32").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1)
conSaveAs.Click();
}
Sys.HighlightObject(pdf);
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    }
ValidationUtils.verify(true,true,"Print Client Invoice is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");
ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet("PDF Invoice",EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")  
  ExcelUtils.WriteExcelSheet("Invoice preparation No",EnvParams.Opco,"Data Management",Invoice_Number)
  ExcelUtils.WriteExcelSheet("Client Invoice No",EnvParams.Opco,"Data Management",Invoice_Number)
  ExcelUtils.WriteExcelSheet("Invoice preparation Job",EnvParams.Opco,"Data Management",Job_Number)
  TextUtils.writeLog("Client Invoice No: "+Invoice_Number);
      }
      
      
      
function validate_Maconomy_Status(){ 

goto_Jobs();
Print_Invoice();
}