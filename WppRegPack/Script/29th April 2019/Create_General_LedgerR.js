﻿//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//var excelName = EnvParams.getEnvironment();


function CreateGeneralR() {
    GotoMenuItem();
    gotoGLTransactions();
    closeAllWorkspaces();
}



var Arrays = [];
var STIME = "";
var Company_no=ExcelUtils.getRowData("Company_no");
var Description=ExcelUtils.getRowData("Description");
var Job_no=ExcelUtils.getRowData("Job_no");
var Work_code1=ExcelUtils.getRowData("Work_code1");
var Amount=ExcelUtils.getRowData("Amount");
var Vendor_no=ExcelUtils.getRowData("Vendor_no");
var Jobno=ExcelUtils.getRowData("Jobno");
var Work_code2=ExcelUtils.getRowData("Work_code2");
var GRP=ExcelUtils.getRowData("GRP");
var Approver = ExcelUtils.getRowData("Approver Name");
var Password = ExcelUtils.getRowData("Approver Password");


function gotoGLTransactions(){
          
          ReportUtils.logStep("INFO", "Create Genreal R started::"+STIME);
          var general = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
          Sys.HighlightObject(general);
          var subgeneral = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
          Sys.HighlightObject(subgeneral);
          /////Ctrl+F////Closefilter///
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Delay(1000);
          var create = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
          Sys.HighlightObject(create);
          create.Click();
          Delay(2000);
          var grid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "");

          
          var company = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
          if(Company_no!=""){
            company.Click();
            Delay(1000);
            WorkspaceUtils.SearchByValue(company,"Company",Company_no)
            }
          else{
            ValidationUtils.verify(false,true,"Company Number is Needed to Create a Journal");    
          }
            
           
          
          Delay(2000);          
          var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
          Sys.HighlightObject(save);
          save.Click();
          Log.Message("Company has Saved");
          Delay(1000);

          var Journal = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "");
          Sys.HighlightObject(Journal);
          var Journal_no = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2).getText().OleValue;
          Log.Message(Journal_no);         


          var entries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
          Sys.HighlightObject(entries);
          var add = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
          add.Click();
          Delay(2000);
          var addgrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
          var cell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McDatePickerWidget", "");
          cell.Click();
          cell.Keys("[Tab]");
          var des = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
          Delay(1000);
          des.Keys("[Tab][Tab][Tab][Tab]");
          
          
          var job = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
//           var companyName = Sys.Process("Maconomy").SWTObject("Shell", "Create Job").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").WaitSWTObject("McValuePickerWidget", "", 2,60000);
          if(Job_no!=""){
          job.Click();
          Delay(1000);
          WorkspaceUtils.SearchByValues(job,"Job",Job_no);
          
          }
          else{ 
            ValidationUtils.verify(false,true,"Job Number is Needed to Create a Entries");
          } 
          
          Delay(4000);
          job.Keys("[Tab]");
          
          var workcode1 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
           if(Work_code1!=""){
            workcode1.Click()
            Delay(1000);
            WorkspaceUtils.SearchByValue(workcode1,"Work Code",Work_code1);
          } 
          else{
            ValidationUtils.verify(false,true,"Work Code is Needed to Create a Entries");
          }
             
          Delay(4000)
          workcode1.Keys("[Tab][Tab]");        
          Delay(1000);
          
          var Debit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
          Debit.setText(Amount); 
          Delay(2000); 
          var entrysave = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
          Sys.HighlightObject(entrysave);
          entrysave.Click();
          Delay(1000);     
           
               
          var Debitedvalue = Debit.getText().OleValue;
          Log.Message("Debited Amount : "+Debitedvalue);
          Delay(2000);
          
          ///////Populated Line ////////
          
          var newadd = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
          newadd.Click();
          Delay(2000);
          var newcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McDatePickerWidget", "", 1);
          newcell.Click();
          newcell.Keys("[Tab][Tab]"); 

          var grp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 3);
          if(GRP!=""){
            grp.Click();
//            WorkspaceUtils.DropDownList(GRP);
            var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000)
            var visible = true;
            if(list.isEnabled()){
              for(var i=1;i<list.getItemCount();i++){
                if(list.getItem(i).getText_2(0)!=null){
                  if(list.getItem(i).getText_2(0).OleValue.toString()==GRP){
                    list.Keys("[Enter]");
                    Delay(5000);
                    break;
                  } 
                  else{ 
                list.Keys("[Down]");
              }        
                } 
                else{ 
                list.Keys("[Down]");
              }
              } 
            } 
          } 
          else{
            ValidationUtils.verify(false,true,"GRP is Needed to Create a Entries");
          } 
          
          Delay(3000);
          grp.Keys("[Tab]");          
          
          
          var vendorno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
          if(Vendor_no!=""){
          vendorno.Click();
          WorkspaceUtils.SearchByValuePicker(vendorno,"Vendor",Vendor_no)
          
          }
          else{ 
            ValidationUtils.verify(false,true,"Vendor Number is Needed to Create a Entries");
          }           
          
          Delay(2000);
          vendorno.Keys("[Tab][Tab]");
          
          
          var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
          if(Jobno!=""){
          jobno.Click();
          Delay(1000);
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x47);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x47);
//          Delay(3000);          
//          Sys.Desktop.KeyDown(0x09);
//          Sys.Desktop.KeyUp(0x09);
          Delay(1000);
          var code = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
          Sys.HighlightObject(code);
          code.setText(Jobno);
          Delay(3000);
          var serch = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", "Search ");
          Sys.HighlightObject(serch);
          serch.Click();
          Delay(5000);
          var table = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
          Sys.HighlightObject(table);
          var itemCount = table.getItemCount();
//          Log.Message(itemCount);
          if(itemCount>0){ 
          for(var i=0;i<itemCount;i++){
            if(table.getItem(i).getText_2(1).OleValue.toString().trim()==Jobno){ 
             var OK = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
                OK.Click();
          
            }
            else{ 
              Sys.Desktop.KeyDown(0x28);
              Sys.Desktop.KeyUp(0x28);
              if(i==itemCount-1){ 
                var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel") ;
                cancel.Click();
                Delay(1000);
                companyName.setText("");
              }
            }
      
            }
          }
          else { 
            var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel"); 
            cancel.Click();
            Delay(1000);
            companyName.setText("");
          } 
          }
          else{ 
            ValidationUtils.verify(false,true,"Job Number is Needed to Create a Entries");
          }  
                         
          Delay(2000);          
          jobno.Keys("[Tab]");
          
          
          var workcode2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
          if(Work_code2!=""){
            workcode2.Click()
            WorkspaceUtils.SearchByValue(workcode2,"Work Code",Work_code2);
          } 
          else{
            ValidationUtils.verify(false,true,"Job Number is Needed to Create a Entries");
          } 
          
          
//          workcode2.setText(Work_code2);
//          Delay(5000);
          workcode2.Keys("[Tab][Tab][Tab]");
          Delay(1000);
          
          
          var Credit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
          Credit.Click(); 
//          Credit.setText(credit);
          Delay(1000);
        
          var populatesave = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
          populatesave.Click();
          Delay(1000);

           
          var attach = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 9);
          attach.Click();
          Delay(5000);
          var upload = Sys.Process("Maconomy").Window("#32770", "Open file", 1);
          Sys.HighlightObject(upload);
          var uploadbar = Sys.Process("Maconomy").Window("#32770", "Open file", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1);
          uploadbar.Keys("C:\\Users\\741496\\Desktop\\test.xlsx");
          Sys.Desktop.KeyDown(0x0D);
          Sys.Desktop.KeyUp(0x0D); 
          Delay(5000);
          Log.Checkpoint("New Document is Attached");
          Delay(1000);
          var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 11);
          Sys.HighlightObject(submit);
          submit.Click();  
          Delay(3000);
          
          var popup = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal")
            for(i=1;i<popup.ChildCount;i++){
               try{
                 var test = popup.Child(i).JavaFullClassName;
                 Log.Message(test);
                 if(test.indexOf('Text')!=-1){
                          Log.Error(Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Text", "").getText())
                          Log.Error("Balance Must be Zero");
                 }
                 if(test.indexOf('Label')!=-1){
                          var ok = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
                          ok.Click();
                          Log.Message("Balance is Zero");
                          Log.Message("Journal is Submitted");
                 }
               }   
               catch(e){
                 Log.Message("error caught while getting text message from popup.");
                }
            } 
          
          var Creditvalue = Credit.getText().OleValue;
          Log.Message("Credited Amount:"+Creditvalue );
          
           if(Debitedvalue == Creditvalue){
                  Log.Message("Debited Amount is Same");
                  Log.Message("Balance is Zero");
                } 
                else
                {
                  Log.Error("Debited Amount is Differnt");                  
                }           
                     
          Delay(1000);
          
          WorkspaceUtils.Rests(Approver,Password);
         
                           
          var todo = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
          Sys.HighlightObject(todo);
          todo.DblClick();
          Delay(3000);
          Sys.Desktop.KeyDown(0x12);
          Sys.Desktop.KeyDown(0x20);
          Delay(2000);
          Sys.Desktop.KeyUp(0x12);
          Sys.Desktop.KeyUp(0x20);
          Delay(2000);
          Sys.Desktop.KeyDown(0x28);
          Sys.Desktop.KeyUp(0x28);
          Delay(2000);
          Sys.Desktop.KeyDown(0x28);
          Sys.Desktop.KeyUp(0x28);
          Delay(2000);
          Sys.Desktop.KeyDown(0x28);
          Sys.Desktop.KeyUp(0x28);
          Delay(2000);
          Sys.Desktop.KeyDown(0x28);
          Sys.Desktop.KeyUp(0x28);
          Delay(2000);
          Sys.Desktop.KeyDown(0x58);
          Sys.Desktop.KeyUp(0x58);  
          Delay(4000);
           
           
           var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
          var refresh;
        Log.Message(childCC)
        for(var i=1;i<=childCC;i++){ 
        refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i)
        if(refresh.isVisible()){ 
        refresh = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("SingleToolItemControl", "");
        refresh.Click();
          Delay(15000);
        Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "")
        if(Client_Managt.isVisible()){ 
        Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", i).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Tree", "");
        Client_Managt.DblClickItem("|GL Journals Awaiting Approval*");
        break;
        }
        }
        }
                 
          
          var general = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
          Sys.HighlightObject(general);
          Delay(3000);
          var generalgrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
          Sys.HighlightObject(generalgrid);
          var companyno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
          Sys.HighlightObject(companyno);
          companyno.Click();
          companyno.Keys("[Tab]");
          var journalno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
          journalno.setText(Journal_no);
          Delay(2000);
          
          Delay(3000)
          var flag=false;
          for(var v=0;v<generalgrid.getItemCount();v++){ 
          if(generalgrid.getItem(v).getText_2(3).OleValue.toString().trim()==Journal_no){ ///column
            flag=true;
            break;
          }
          else{ 
            generalgrid.Keys("[Down]");
          }
         }

         ValidationUtils.verify(false,true,"Journal is Created and available in Maconomy"); 
          
//          var close = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");
//          close.Click();
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
          
          Delay(1000);
          var post = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 11);
          post.Click();    
          Delay(2000);
          ValidationUtils.verify(false,true,"Journal is Posted");
          
          var postpopup = Sys.Process("Maconomy").SWTObject("Shell", "GL Journals Awaiting Approval - General Journal");
          Sys.HighlightObject(postpopup);
          var postconform = Sys.Process("Maconomy").SWTObject("Shell", "GL Journals Awaiting Approval - General Journal").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
          postconform.Click();   
               

} 




  function GotoMenuItem(){
    var menubar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    menubar.DblClick();
      if(ImageRepository.ImageSet4.GL.Exists()){
        ImageRepository.ImageSet4.GL.Click();
      } 
      else if(ImageRepository.ImageSet4.GL1.Exists()){
        ImageRepository.ImageSet4.GL1.Click();
      } 
      else{
        ImageRepository.ImageSet4.GLs.Click();
      } 
  
        var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
        var General;
       for(var i=1;i<=childCC;i++){ 
          General = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
          if(General.isVisible()){ 
          General = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
          General.DblClickItem("|GL Transactions");
       }
      }
     Delay(6000);
  } 

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Ctrl
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); //Ctrl
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}







