//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//var excelName = EnvParams.getEnvironment();


var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "Post Journal Entries", true);
//Client Open Statement


function ClientOpen() {
    GotoMenuItem();
    gotoLookup();
    closeAllWorkspaces();
}

var STIME = "";
var checkmark = false;
var Arrays = [];
var Company_no= ExcelUtils.getRowData("Company_no");
Log.Message(Company_no)
//var Client_no= ExcelUtils.getRowData("Client_no");
var Journal_no= ExcelUtils.getRowData("Journal_no");
Log.Message(Journal_no)


function gotoLookup(){
      ReportUtils.logStep("INFO", "Client Open Statement is started::"+STIME);
      var periodic = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5);
      Sys.HighlightObject(periodic);
      periodic.Click();
      Delay(2000);
      var clientgrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
      Sys.HighlightObject(clientgrid);
      var clientcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 4).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
      clientcell.Click();
      clientcell.setText(Client_no);
      
      Delay(3000)
      var flag=false;
      for(var v=0;v<clientgrid.getItemCount();v++){ 
        if(clientgrid.getItem(v).getText_2(0).OleValue.toString().trim()==Client_no){ 
          flag=true;
          break;
        }
        else{ 
          clientgrid.Keys("[Down]");
        }
       }

       ValidationUtils.verify(flag,true,"Client Number is available in Maconomy"); 
      
  
      Sys.Desktop.KeyDown(0x11);
      Sys.Desktop.KeyDown(0x46);
      Sys.Desktop.KeyUp(0x11);      
      Sys.Desktop.KeyUp(0x46);
      Delay(1000);
      var openentry = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
      Sys.HighlightObject(openentry);
      var  printicon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 5);
      Sys.HighlightObject(printicon);
      printicon.Click();
      Delay(2000);

      
      var printgrid = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement") ;
      Sys.HighlightObject(printgrid);
      Delay(2000);
      var client = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
      if(Client_no!=""){
        client.Click();
        WorkspaceUtils.SearchByValuePicker(client,"Client",Client_no)
      } 
      else{
        ValidationUtils.verify(false,true,"Client Number is Needed to Print");
      } 
      Delay(1000);
      Sys.Desktop.KeyDown(0x09);
      var clientno = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 4);
      if(Client_no!=""){
        clientno.Click();
        WorkspaceUtils.SearchByValuePicker(clientno,"Client",Client_no)
      } 
      else{
        ValidationUtils.verify(false,true,"Client Number is Needed to Print");
      } 
      
      Delay(1000);
      var company = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
      if(Company_no!=""){
        company.Click();
        WorkspaceUtils.SearchByValue(company,"Company",Company_no)
      } 
      else{
        ValidationUtils.verify(false,true,"Company Number is Needed to Print");
      } 
      
      Delay(1000);
      var companyno = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 4)
      if(Company_no!=""){
        companyno.Click();
        WorkspaceUtils.SearchByValue(companyno,"Company",Company_no)
      } 
      else{
        ValidationUtils.verify(false,true,"Company Number is Needed to Print");
      } 
      
      Delay(1000);
      var zerobal = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 14).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
      zerobal.Click();
      var print = Sys.Process("Maconomy").SWTObject("Shell", "Print Client Open Entry Statement").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Print")
      print.Click();
//      Delay(1000);
      ValidationUtils.verify(true,true,"Client Statement is Printed");
      
      var exit = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - SSC IN -  SSC Combined Biller, Ics").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);     
      exit.Click();   
}


 function GotoMenuItem(){
    var menubar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    menubar.DblClick();
      if(ImageRepository.ImageSet6.AccRecee.Exists()){
        ImageRepository.ImageSet6.AccRecee.Click();
      } 
      else if(ImageRepository.ImageSet6.AccRec1.Exists()){
        ImageRepository.ImageSet6.AccRec1.Click();
      } 
      else{
        ImageRepository.ImageSet6.AccRec.Click();
      }   
      
      
      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
      var Account_Recevable;
       for(var i=1;i<=childCC;i++){ 
          Account_Recevable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
          if(Account_Recevable.isVisible()){ 
          Account_Recevable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
          Account_Recevable.DblClickItem("|AR Lookups");
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







 