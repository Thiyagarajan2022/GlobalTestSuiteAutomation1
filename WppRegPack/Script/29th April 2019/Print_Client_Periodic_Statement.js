﻿//USEUNIT WorkspaceUtils
//USEUNIT ExcelUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//var excelName = EnvParams.getEnvironment();

var excelName = EnvParams.getEnvironment();
ExcelUtils.setExcelName(Project.Path+excelName, "Periodic Statement", true);



function Periodic() {
    GotoMenuItem();
    gotoLookup();
    closeAllWorkspaces();
}


var STIME = "";
var Arrays = [];
var Company_no= ExcelUtils.getRowData("Company_no");
var Client_no= ExcelUtils.getRowData("Client_no");



function gotoLookup(){
      ReportUtils.logStep("INFO", "Print Client Periodic is started::"+STIME);
      var periodic = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
      Sys.HighlightObject(periodic);
      periodic.Click();
      var clientgrid = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
      Sys.HighlightObject(clientgrid);
      var clientcell = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
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
      Delay(2000);
      var printicon = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
      printicon.Click();
      Delay(2000);
      var printgrid = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement");
      Sys.HighlightObject(printgrid);
      Delay(1000);
      
      var client = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 2);
      if(Client_no!=""){
        client.Click();
        WorkspaceUtils.SearchByValuePicker(client,"Client",Client_no);
      } 
      else{
        ValidationUtils.verify(false,true,"Client Number is Needed to Print");
      } 
      
      Delay(2000);
      Sys.Desktop.KeyDown(0x09);
      var clientno = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 1).SWTObject("McValuePickerWidget", "", 4);
      if(Client_no!=""){
        clientno.Click();
        WorkspaceUtils.SearchByValuePicker(clientno,"Client",Client_no);
      } 
      else{
        ValidationUtils.verify(false,true,"Client Number is Needed to Print");
      } 
      var company = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 2);
      if(Company_no!=""){
        company.Click();
        WorkspaceUtils.SearchByValue(company,"Company",Company_no);
      } 
      else{
        ValidationUtils.verify(false,true,"Company Number is Needed to Print");
      } 
      Sys.Desktop.KeyDown(0x09);
      var companyno = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 2).SWTObject("McValuePickerWidget", "", 4);
      if(Company_no!=""){
        companyno.Click();
        WorkspaceUtils.SearchByValue(companyno,"Company",Company_no);
      } 
      else{
        ValidationUtils.verify(false,true,"Company Number is Needed to Print");
      } 
      var zerobal = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("McGroupWidget", "").SWTObject("Composite", "", 15).SWTObject("McPlainCheckboxView", "", 2).SWTObject("Button", "");
      zerobal.Click();
      Delay(1000);
      var print = Sys.Process("Maconomy").SWTObject("Shell", "Print Periodic Client Statement").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Print");
      print.Click();  
      ValidationUtils.verify(true,true,"Journal No is Printed");
      
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
      
      
      var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
      var Account_Recevable;
       for(var i=1;i<=childCC;i++){ 
          Account_Recevable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
          if(Account_Recevable.isVisible()){ 
          Account_Recevable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
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






