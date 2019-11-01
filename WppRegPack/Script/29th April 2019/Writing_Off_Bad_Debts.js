//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//var excelName = EnvParams.getEnvironment();


var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "Bad Debts", true);




function BadDebts() {
    GotoMenuItem();
    gotoGLTransactions();
    closeAllWorkspaces();
}

var STIME = "";
var Arrays = [];
var Company_no= ExcelUtils.getRowData("Company_no");
var Description= ExcelUtils.getRowData("Description");
var GRP= ExcelUtils.getRowData("GRP");
Log.Message(GRP)
var Job_no= ExcelUtils.getRowData("Job_no");
var Work_code1= ExcelUtils.getRowData("Work_code1");
var Amount= ExcelUtils.getRowData("Amount"); 
var Vendor_no= ExcelUtils.getRowData("Vendor_no");
var Jobno= ExcelUtils.getRowData("Jobno");
var Work_code2= ExcelUtils.getRowData("Work_code2");
var Client_no= ExcelUtils.getRowData("Client_no");
Log.Message(Client_no)
var Depart= ExcelUtils.getRowData("Depart");
var Unit= ExcelUtils.getRowData("Unit");
var GRP1= ExcelUtils.getRowData("GRP1");



function gotoGLTransactions(){    
          ReportUtils.logStep("INFO", "Bad Debts is started::"+STIME);  
          var general = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
          Sys.HighlightObject(general);
          var subgeneral = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
          Sys.HighlightObject(subgeneral);
          /////Ctrl+F////Closefilter///
          Sys.Desktop.KeyDown(0x11);          
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);          
          Sys.Desktop.KeyUp(0x46);
          var create = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
          Sys.HighlightObject(create);
          create.Click();
          Delay(2000);
          var grid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "");
//          Sys.HighlightObject(grid);
          var company = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McValuePickerWidget", "", 3);
          if(Company_no!=""){
            company.Click();
            WorkspaceUtils.SearchByValue(company,"Company",Company_no);
          }
          else{
            ValidationUtils.verify(false,true,"Company Number is Needed to Create a Journal");
          } 
          
          Delay(1000);
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
          cell.Keys("[Tab][Tab]");
          
          //////dropdown///
          
          var grpcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 3);
//          grpcell.Keys("G");
          Delay(4000);
          if(GRP!=""){
          grpcell.Click();
          var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,6000)
          for(var i=1;i<list.getItemCount();i++){ 
              if(list.getItem(i).getText_2(0)!=null){
                if(list.getItem(i).getText_2(0).OleValue.toString().trim()==GRP){
                  list.Keys("[Enter]");
                  Delay(5000);
                  break;
                }
                else{ 
                    list.Keys("[Down]");
                    Delay(5000);
                }  
              } 
               else{ 
                  list.Keys("[Down]"); 
                  Delay(4000);         
            }
          }
          }
          else{ 
              ValidationUtils.verify(false,true,"GRP is Needed to Create a Entries");
          }

          Delay(2000);
          grpcell.Keys("[Tab]");
          
         var clientcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
          if(Client_no!=""){
            clientcell.Click();              
            Delay(1000);
            Sys.Desktop.KeyDown(0x11);
            Sys.Desktop.KeyDown(0x47);
            Sys.Desktop.KeyUp(0x11);
            Sys.Desktop.KeyUp(0x47);
            Delay(3000);
            var code = Sys.Process("Maconomy").SWTObject("Shell", "Client").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
            code.setText(Client_no);
            Delay(3000);
            var serch = Sys.Process("Maconomy").SWTObject("Shell", "Client").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", "Search ")
            Sys.HighlightObject(serch);
            serch.Click();
            Delay(5000);
            var table = Sys.Process("Maconomy").SWTObject("Shell", "Client").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
            Sys.HighlightObject(table);
            var itemCount = table.getItemCount();
            if(itemCount>0){ 
            for(var i=0;i<itemCount;i++){
              if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Client_no){ 
               var OK = Sys.Process("Maconomy").SWTObject("Shell", "Client").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
                  OK.Click();
          
              }
              else{ 
                Sys.Desktop.KeyDown(0x28);
                Sys.Desktop.KeyUp(0x28);
                if(i==itemCount-1){ 
                  var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Client").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel")
                  cancel.Click();
                  Delay(1000);
                  clientcell.setText("");
                }
              }
      
              }
            }
            else { 
              var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Client").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel")
              cancel.Click();
              Delay(1000);
              clientcell.setText("");
            }          
          
          } 
          else{
            ValidationUtils.verify(false,true,"Client Number is Needed to Create a Entries");
          } 

         Delay(3000);
         clientcell.Keys("[Tab][Tab][Tab][Tab][Tab][Tab]");
          
          var Credit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
          Credit.Click();
          Credit.setText(Amount);
          Credit.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]")
          Delay(2000); 
          
         var  depart = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
         if(Depart!=""){
           depart.Click();
           WorkspaceUtils.SearchByValue(depart,"Local Specification 2",Depart)
         } 
         else{
            ValidationUtils.verify(false,true,"Department is Needed to Create a Entries");
          }          
         Delay(1000);
         depart.Keys("[Tab][Tab]");
         
         var unit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
         if(Unit!=""){
           unit.Click();
           WorkspaceUtils.SearchByValue(unit,"Local Specification 4",Unit)
         } 
         else{
            ValidationUtils.verify(false,true,"Unit is Needed to Create a Entries");         
         } 
         
         Delay(2000);
          
          var entrysave = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
          Sys.HighlightObject(entrysave);
          entrysave.Click();
          Delay(1000);     
           
          ///////Populated Line ////////
          
          var newadd = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
          newadd.Click();
          Delay(2000);
          var newcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McDatePickerWidget", "", 1);
          newcell.Click();
          newcell.Keys("[Tab][Tab]"); 
          
          var GRP = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "", 3);
          GRP.Keys("G");
//          Delay(4000);
//          if(GRP1!=""){
//          GRP.Click();
//          var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,6000)
//          for(var i=1;i<list.getItemCount();i++){ 
//              if(list.getItem(i).getText_2(0)!=null){
//                if(list.getItem(i).getText_2(0).OleValue.toString().trim()==GRP1){
//                  list.Keys("[Enter]");
//                  Delay(5000);
//                  break;
//                }
//                else{ 
//                    list.Keys("[Down]");
//                    Delay(1000);
//                }  
//              } 
//               else{ 
//                  list.Keys("[Down]"); 
//                  Delay(1000);         
//            }
//          }
//          }
//          else{ 
//              ValidationUtils.verify(false,true,"GRP is Needed to Create a Entries");
//          }
          
          GRP.Keys("[Tab]");
          Delay(1000);
          
          var vendorno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
          if(Vendor_no!=""){
            vendorno.Click();
            Delay(1000);
            Sys.Desktop.KeyDown(0x11);
            Sys.Desktop.KeyDown(0x47);
            Sys.Desktop.KeyUp(0x11);
            Sys.Desktop.KeyUp(0x47);
            Delay(3000);
            var code = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
            code.setText(Client_no);
            Delay(3000);
            var serch = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", "Search ");
            Sys.HighlightObject(serch);
            serch.Click();
            Delay(5000);
            var table = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
            Sys.HighlightObject(table);
            var itemCount = table.getItemCount();
            if(itemCount>0){ 
            for(var i=0;i<itemCount;i++){
              if(table.getItem(i).getText_2(0).OleValue.toString().trim()==Vendor_no){ 
               var OK = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
                  OK.Click();
          
              }
              else{ 
                Sys.Desktop.KeyDown(0x28);
                Sys.Desktop.KeyUp(0x28);
                if(i==itemCount-1){ 
                  var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
                  cancel.Click();
                  Delay(1000);
                  vendorno.setText("");
                }
              }
      
              }
            }
            else { 
              var cancel = Sys.Process("Maconomy").SWTObject("Shell", "Vendor").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
              cancel.Click();
              Delay(1000);
              vendorno.setText("");
            }          
          
          } 
          else{
            ValidationUtils.verify(false,true,"Client Number is Needed to Create a Entries");
          } 
          
          Delay(4000);
          vendorno.Keys("[Tab][Tab]");
          
          var jobno = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
            if(Jobno!=""){
              jobno.Click();
              WorkspaceUtils.SearchByValues(jobno,"Job",Jobno)
            } 
            else{
            ValidationUtils.verify(false,true,"Job Number is Needed to Create a Entries");
          } 
          Delay(2000);          
          jobno.Keys("[Tab]");
          
          var workcode2 = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "", 4);
          if(Work_code2!=""){
            workcode2.Click();
            WorkspaceUtils.SearchByValue(workcode2,"Work Code",Work_code2)
          } 
          else{
            ValidationUtils.verify(false,true,"Work Code is Needed to Create a Entries");
          } 
          Delay(5000)       
          
          var populatesave = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
          populatesave.Click();
          Delay(3000);
          Log.Checkpoint("Entries has Saved");
          
          
          var addgrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
          var column = addgrid.getColumnCount();
          Log.Message(column);
          var row = addgrid.getItemCount();
          Log.Message(row);
          let x=0;y=1;
              var credit = (addgrid.getItem(x).getText_2(9));
              Log.Message(credit);
              var debit = (addgrid.getItem(y).getText_2(8));
              Log.Message(debit);
          
          
          
          var submit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 11);
          Sys.HighlightObject(submit);
          submit.Click();  
          Delay(3000);
          Log.Message("Journal is Submitted");
          
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
          
//          var popup = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal");
//          Sys.HighlightObject(popup);
//          
//          if(Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Label", "*").getText().OleValue.toString().indexOf("Journal no.")==-1)
//          {
//            var ok = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
//            ok.Click(); 
//          } 
//          else
//          {
//          Log.Error(Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Text", "").getText())
//          Log.Error("Balance Must be Zero");
//          }
//          
            if(debit = credit){
                  Log.Message("Debited Amount is Same");
                  Log.Message("Balance is Zero");
                } 
                else{
                  Log.Error("Debited Amount is Differnt");                  
           }


          
          var macanomy = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *");          
          Sys.HighlightObject(macanomy);
          Sys.Desktop.KeyDown(0x12);
          Sys.Desktop.KeyUp(0x12);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x46);
          Sys.Desktop.KeyDown(0x52);
          Sys.Desktop.KeyUp(0x52);                              
          Delay(22000);
          
          var username = "SSC IN -  SSC Combined Biller, Ics";                                //Macanomy Login  level zero
          var pwd = "CORE@WPP123";
          Delay(24000);
          var userName = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1); 
          userName.SetFocus();
          userName.setText("^a[BS]");
          userName.setText(username);
          Delay(3000);
          var pwdword = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2); 
          pwdword.SetFocus();
          pwdword.setText(pwd); 
          Delay(1000)
          var loginbtn = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
          Sys.HighlightObject(loginbtn);
          loginbtn.Click();
          Delay(8000);
          
          GotoMenuItem();
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
          var flag=false;
          for(var v=0;v<generalgrid.getItemCount();v++){ 
            if(generalgrid.getItem(v).getText_2(2).OleValue.toString().trim()==Journal_no){ 
              flag=true;
              break;
            }
            else{ 
              generalgrid.Keys("[Down]");
            }
           }

         ValidationUtils.verify(flag,true,"Journal Number is Created and available in Maconomy");           
          
          
          Sys.Desktop.KeyDown(0x11);
          Sys.Desktop.KeyDown(0x46);
          Sys.Desktop.KeyUp(0x11);
          Sys.Desktop.KeyUp(0x46);
          Delay(1000);
          var post = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 11);
          post.Click();    
          Delay(2000);
          Log.Message("Journal is Posted"); 
          var postpopup = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal")
          Sys.HighlightObject(postpopup);
          var postconform = Sys.Process("Maconomy").SWTObject("Shell", "GL Transactions - General Journal").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
          postconform.Click();   
            

} 



  function GotoMenuItem(){
    var menubar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    menubar.DblClick();
      if(ImageRepository.ImageSet4.GLs.Exists()){
        ImageRepository.ImageSet4.GLs.Click();
      } 
      else if(ImageRepository.ImageSet4.GL.Exists()){
        ImageRepository.ImageSet4.GL.Click();
      } 
      else{
        ImageRepository.ImageSet4.GL1.Click();
      } 
      
           var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
           var Bad_Debt;
          for(var i=1;i<=childCC;i++){ 
          Bad_Debt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
          if(Bad_Debt.isVisible()){ 
          Bad_Debt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
          Bad_Debt.DblClickItem("|GL Transactions");
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



