//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT ReportUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams

var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "JobCreation", true);
var Arrays = [];
var count = true;
var checkmark = false;
var STIME = "";
var comapany = ExcelUtils.getRowData("comapany")
var Job_group =ExcelUtils.getRowData("Job_group")
var Job_Type =ExcelUtils.getRowData("Job_Type")
var department =ExcelUtils.getRowData("department")
var buss_unit = ExcelUtils.getRowData("buss_unit")
var TemplateNo =ExcelUtils.getRowData("TemplateNo")
var Product =ExcelUtils.getRowData("Product")
var Job_name=ExcelUtils.getRowData("Job_name")
var Project_manager =ExcelUtils.getRowData("Project_manager")



function createAJob() {
ReportUtils.logStep("INFO", "Entering job details");
  Delay(6000);
  var all_job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
  all_job.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    
  var newJobBtn = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  newJobBtn.Click();
  Delay(1000);
 
  var companyName = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,60000);
  if(comapany!=""){
  companyName.Click();
  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(comapany);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==comapany){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        companyName.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Company").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    companyName.setText("");
  } 
  }
  else{ 
    ValidationUtils.verify(false,true,"Company Number is Needed to Create a Job");
  }
    
  
    
    
    


  var job = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McPopupPickerWidget", "", 2);

if(Job_group!=""){
job.Click();
var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
var Add_Visible2 = true;
while(Add_Visible2){
if(list.isEnabled()){
Add_Visible2 = false;
    for(var i=0;i<list.getItemCount();i++){ 
      if(list.getItem(i).getText_2(0)!=null){ 
        if(list.getItem(i).getText_2(0).OleValue.toString().trim()==Job_group){ 
          list.Keys("[Enter]");
          Delay(5000);
          break;
        }else{ 
          list.Keys("[Down]");
        }
          
      }else{ 
        list.Keys("[Down]");
      }
    }
}
}
}
else{ 
    ValidationUtils.verify(false,true,"Company Group is Needed to Create a Job");
  }
  
   
  var JobType = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(Job_Type!=""){
  JobType.Click();
  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Job Type").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(Job_Type);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Job Type").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Job Type").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Job_Type){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Job Type").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Job Type").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        JobType.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Job Type").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    JobType.setText("");
  }
    }
else{ 
    ValidationUtils.verify(false,true,"Company Type is Needed to Create a Job");
  }

    
    
    
    
  var Depart = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
 if(department!=""){
  Depart.Click();
  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Department").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(department);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Department").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Department").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==department){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Department").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Department").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        Depart.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Department").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    Depart.setText("");
  }
        }
else{ 
    ValidationUtils.verify(false,true,"Department Number is Needed to Create a Job");
    }  
 
    
    
    
    
  var BussUnit = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  if(buss_unit!=""){
  BussUnit.Click();
  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(buss_unit);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==buss_unit){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        BussUnit.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Business Unit").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    BussUnit.setText("");
  } 
            }
else{ 
    ValidationUtils.verify(false,true,"Business Unit Number is Needed to Create a Job");
    }  
    
    
    
    
  var template = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(TemplateNo!=""){
  template.Click();
  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(TemplateNo);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==TemplateNo){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        template.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    template.setText("");
  }
        }
else{ 
    ValidationUtils.verify(false,true,"Templete Number is Needed to Create a Job");
    }   
    
    
    
    
    
    
    
    
    
  var prdNumber = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
if(Product!=""){
  prdNumber.Click();


  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Product Result").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  Delay(1000);
  code.Keys("[Tab]");
  Delay(1000);
  code.setText(Product);
    
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Product Result").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Product Result").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);

  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(1).OleValue.toString().trim()==Product){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Product Result").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Product Result").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        prdNumber.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Product Result").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    prdNumber.setText("");
  }
      }
else{ 
    ValidationUtils.verify(false,true,"Product Number is Needed to Create a Job");
    } 

    
    
    
    
    
    
    
  var jobName = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 8).SWTObject("McTextWidget", "", 2);
  jobName.setText(Job_name+" "+STIME);
  Delay(5000);

 if(Project_manager!=""){
  var ProjectManger = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
  ProjectManger.Click();
  if(ProjectManger.getText()!=Project_manager){
  Delay(1000);
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  Delay(3000);
  var code = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(Project_manager);
  Delay(3000);
  var serch = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
  Sys.HighlightObject(serch);
  serch.Click();
  Delay(5000);
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Project_manager){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK")
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
        cancel.Click();
        Delay(1000);
        ProjectManger.setText("");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Employee").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
    cancel.Click();
    Delay(1000);
    ProjectManger.setText("");
  }
  } 
  }
    
var btnCreate = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Create");    
if(btnCreate.isEnabled()){
Log.Message("Create Button is Vissible");
Log.Message("Job is CREATED");
  Delay(1000);
  Sys.HighlightObject(btnCreate)
  btnCreate.Click();
ReportUtils.logStep("INFO", Job_name+" "+STIME +" : is Created");
}
else{ 
  Delay(4000);
  var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Button", "Cancel");
  Delay(1000);
  Sys.HighlightObject(cancel)
  cancel.Click();
ReportUtils.logStep("ERROR", "Job is not Created");
}
    
  Delay(4000);
  
}

function GoToJob() {
  var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  Delay(2000);

  var companyFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").
  SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).
  SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).
  SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).
  SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").
  SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  companyFilter.forceFocus();
  companyFilter.setVisible(true);
  companyFilter.ClickM();
  table.Child(0).setText("^a[BS]");
  table.Child(0).setText(comapany);
  Delay(1000);
  Sys.Desktop.KeyDown(0x09); // Press Ctrl
  Delay(1000);
  Sys.Desktop.KeyDown(0x09);
  Delay(1000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  Sys.Desktop.KeyUp(0x09);
  var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 3);
  job.Click();

  job.setText(Job_name+" "+STIME);
  Delay(3000);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==(Job_name+" "+STIME)){ 
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

  ValidationUtils.verify(flag,true,"Job Created is available in system");
  if(flag){
  ReportUtils.logStep("INFO", "Created Job is listed in table");
  closeFilter.Click();
  Delay(8000);
//    var Home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//    Home.Click();
  if(count){
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
  info.Click();
  count=false;
  }
  Delay(3000);
  var Templete_Job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 3).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Blanket_invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 4).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "")
  var estimation = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 8).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Amount_Registrations = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 6).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  var Invocing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 7).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", ""); 
  var TimeReg = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 3).SWTObject("Composite", "", 5).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
  checkmark = false;
  if(Templete_Job.getSelection()){ 
  Templete_Job.Click();
    Log.Message("Templete_Job is UnChecked")
    checkmark = true;
  }
  if(Blanket_invoice.getSelection()){ 
  Blanket_invoice.Click();
    Log.Message("Blanket_invoice is UnChecked")
    checkmark = true;
  }
  if(Amount_Registrations.getSelection()){ 
  Amount_Registrations.Click();
    Log.Message("Amount_Registrations is UnChecked")
    checkmark = true;
  }
  if(Invocing.getSelection()){ 
  Invocing.Click();
    Log.Message("Invocing is UnChecked")
    checkmark = true;
  }
  if(TimeReg.getSelection()){ 
  TimeReg.Click();
    Log.Message("TimeRegistration is UnChecked")
    checkmark = true;
  }
  if(estimation.getSelection()){ 
  estimation.Click();
    Log.Message("Estimating is UnChecked")
    checkmark = true;
  }

  if(checkmark){ 
      
    Delay(3000);
    var save_change = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 4)
    save_change.Click();
    Log.Message("Changes is Saved");
    Delay(5000);
  }
  ReportUtils.logStep("INFO", "Job is Saved Pending for Approval");
  var filter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("SingleToolItemControl", "", 2);
  filter.Click();
}
}
function goToJobMenuItem(){

//   var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "");
  var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  menuBar.DblClick();
 if(ImageRepository.ImageSet.Jobs.Exists()){
 ImageRepository.ImageSet.Jobs.Click();// GL
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
}
var jobSubItem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 4).SWTObject("Tree", "");
//  var jobSubItem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "")
jobSubItem.DblClickItem("|Jobs"); 
Delay(5000); 
ReportUtils.logStep("INFO", "Moved to Jobs from job Menu");
 
}







function createJob() {

STIME = WorkspaceUtils.StartTime();
ReportUtils.logStep("INFO", "Creat Job test started::"+STIME);
goToJobMenuItem();   
createAJob();   
GoToJob();
WorkspaceUtils.closeAllWorkspaces();
  
}



