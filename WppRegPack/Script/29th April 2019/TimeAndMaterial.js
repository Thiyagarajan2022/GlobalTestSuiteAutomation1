//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
var excelName = EnvParams.getEnvironment();
var sheetName = "Time&MaterialInvoicing";
var JobDetail = [];
var invoices = [];
var approvers ="";
var Approve_Level = [];
var HRData = [];
var LoginEmp = [];
var UserPasswd = [];
var arrys = [];
var checkpoint;
var zz = 1;
var jobName;
var jobNo;
var AAArrays = [];
var uniqueAAray = [];

function SOXexcel(CreateClient,start){ 
var Arrayss = []; 
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];

   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();
     while (!DDT.CurrentDriver.EOF()) {
      
      var temp ="";
       if(xlDriver.Value(colsList[start])!=null){
      temp = temp+xlDriver.Value(start).toString().trim();
      }
      else{ 
        temp = temp;
      }
     Arrayss[id]=temp;
//     Log.Message(temp);
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
return Arrayss;
}

 function excel(CreateClient,start){ 

var Arrayss = [];
var xlDriver = DDT.ExcelDriver(Project.Path+excelName, sheetName, true);
var id =0;
var colsList = [];
Log.Message(DDT.CurrentDriver.ColumnCount);
   for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
     colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
   }
//   xlDriver.Next();

     while (!DDT.CurrentDriver.EOF()) {
     var temp ="";
      for(var idx=start;idx<colsList.length;idx++){  
       if(xlDriver.Value(colsList[idx])!=null){
      temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
      }
      else{ 
        temp = temp+"*";
      }
      }
     if(temp.length!=6){
     Arrayss[id]=temp;
     Log.Message(temp)
     }
     id++;     
     xlDriver.Next();
     }
     DDT.CloseDriver(xlDriver.Name);
     return Arrayss;
}




function TimeMaterialInvoicing(){ 
gotoMenu();
v1();
v2();
v3();
}

function gotoMenu(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();// GL
}
else if(ImageRepository.ImageSet.Jobs.Exists()){
ImageRepository.ImageSet.Jobs.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
}

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|Jobs");
}

}
Delay(6000);
//invoice();
}


function v1(){ 
    var closeFilter = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "");

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  Delay(2000);
JobDetail = SOXexcel(sheetName,1);
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
  table.Child(0).setText(JobDetail[0]);
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

  job.setText(JobDetail[1]);
  Delay(3000);
  var flag=false;
  for(var v=0;v<table.getItemCount();v++){ 
    if(table.getItem(v).getText_2(3).OleValue.toString().trim()==JobDetail[1]){ 
    jobNo = table.getItem(v).getText_2(2).OleValue.toString().trim();
      flag=true;
      break;
    }
    else{ 
      table.Keys("[Down]");
    }
  }

//  ValidationUtils.verify(flag,true,"Job Created is available in system");
  if(flag){
  closeFilter.Click();
  Delay(8000);
    var Home = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 5)
    Home.Click();
//Delay(4000);
  var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3)
  ref.Refresh();
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).isVisible())
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).isVisible())
  var FullBudget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
    
  Sys.HighlightObject(FullBudget)  ;
  FullBudget.Click();
  Delay(5000);
//  ref.Refresh();
  var budgetTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 5).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
  Sys.HighlightObject(budgetTable);
  var show_budget = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 1).SWTObject("McPopupPickerWidget", "", 2);
  Sys.HighlightObject(show_budget);
  show_budget.Keys("Working Estimate");
  Delay(6000);

  if(budgetTable.getItemCount()>1){
  for(var j=0;j<budgetTable.getItemCount()-1;j++){  
    var temp = "";
    for(var i=0;i<budgetTable.getColumnCount();i++){ 
    if((budgetTable.getItem(j).getText_2(i)!="") && ((i==1)||(i==0))){
     temp = temp+ budgetTable.getItem(j).getText_2(i).OleValue.toString().trim()+"*";
     }
     if((budgetTable.getItem(j).getText_2(i)!="") && (i==8)){
     temp = temp+ aqConvert.StrToInt(budgetTable.getItem(j).getText_2(i).OleValue.toString().trim());
     }
    }
    arrys [j] = temp;
      Log.Message(temp);
    }
    }
    ref.Refresh();
    var invoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 8).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 8);
    invoice.Click();  
    Delay(5000);
    ref.Refresh();
    var invoicehistory = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 10);
    invoicehistory.Click();
    Delay(3000);
    ref.Refresh();
    var tabs = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 9).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
//    var AAArrays = [];
//    var uniqueAAray = [];
    var zz =0;
    for(var jj=0;jj<tabs.getItemCount();jj++){ 
    if(jj!=0)
    tabs.Keys("[Down]");
    Delay(4000);
    var specification = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
    specification.Click();  
    Delay(3000);
    var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 - Finance").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table)  ;  
    

    for(var ii=0;ii<table.getItemCount();ii++){ 
      AAArrays[zz] = table.getItem(ii).getText_2(0).OleValue.toString().trim()+"*"+table.getItem(ii).getText_2(3).OleValue.toString().trim();
      zz++
    }
   
    }
    zz =0;
    for(var jj=0;jj<AAArrays.length;jj++){
     Log.Message("All :"+AAArrays[jj]);
    }
}
}


function v2(){ 
    var ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
  ref.Refresh();  
      Delay(3000);
  var info = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 6).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 6);
  info.Click();
  ref.Refresh();
  Delay(3000);
//  invoices = excel(sheetName,2);
  var selectionBilling = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  selectionBilling.Click();
  Delay(5000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);

  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
//============================================================================
zz = AAArrays.length;
    for(var ii=0;ii<table.getItemCount();ii++){ 
var rr = table.getItem(ii).getText_2(1).OleValue.toString().trim();
      AAArrays[zz] = rr.substring(rr.indexOf(" - ")+3)+"*"+table.getItem(ii).getText_2(4).OleValue.toString().trim();
      Log.Message("ALL :"+AAArrays[zz]);
      zz++
    }
      var AArrays = [];
    var uniqueAAray = [];
    for(var ii=0;ii<AAArrays.length;ii++){ 
    var temp = AAArrays[ii].split("*");
      AArrays[ii] = temp[0];
    }
    
    uniqueAAray = Array.from(new Set(AArrays));
    for(var ii=0;ii<uniqueAAray.length;ii++){ 
      Log.Message("Unique :"+uniqueAAray[ii]);
    }    
    
    AArray = [];
    var j=0;
    for(var ij=0;ij<uniqueAAray.length;ij++){
    var temp = "0";
    for(var ii=0;ii<AAArrays.length;ii++){ 
    var temps = AAArrays[ii].split("*");
        if(temps[0]==uniqueAAray[ij]){
        if(temps[1]!=""){
        temp = aqConvert.StrToInt(temp.toString().trim())+aqConvert.StrToInt(temps[1].toString().trim());
        }
        }
    }
    AArray[j] = uniqueAAray[ij]+"*"+temp;
    j++;
    Log.Message(uniqueAAray[ij]+"*"+temp);
    }
 //---------------------   
     invoices = excel(sheetName,2);
    for(var i=0;i<AArray.length;i++){ 
    var temp = AArray[i].split("*");
      for(var j=0;j<arrys.length;j++){ 
         var temp1 = arrys[j].split("*");
         for(var k=0;k<invoices.length;k++){
         var temp2 = invoices[k].split("*");
         if((temp[1] == temp1[2])&&(temp2[0]==temp1[0])) { 
           Log.Message(temp1[0]+ " budget is Fully invoiced");
           Log.Error(temp1[0]+ " budget is Fully invoiced");
         }
         }
      }
    }
    }
    
    
//=============================================================================

  function v3(){
  var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
  var selectionBilling = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  for(var i=0;i<table.getItemCount();i++){
  selectionBilling.Click();
  checkpoint = false;
  for(var j=0;j<invoices.length;j++){
  var temp = invoices[j].split("*");
//  Log.Message("invoices :"+invoices[j])
if(temp[4]!="")

  if((table.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(temp[0])!=-1) && (table.getItem(i).getText_2(10).OleValue.toString().trim().indexOf(temp[4])!=-1)){
  checkpoint = true;
  if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).isVisible()){
  var entries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
  entries.Click();
  }

  if(!Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("PTabItemPanel", "", 3).isVisible()){
  var entries = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
  entries.Click();
  }
  Delay(2000);
  
  var TTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
  
  var add = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);
  Sys.HighlightObject(add);
  add.Click();
  
  var emp = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  if(temp[1]!=""){
  emp.Click();
  WorkspaceUtils.SearchByValue(emp,"Employee",temp[1]);
  }else{ 
  ValidationUtils.verify(false,true,"Employee No is Needed to Invoicing");
  }
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  
  var quantity = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  if(temp[2]!=""){
  quantity.Click();
  quantity.setText(temp[2]);
  }else{ 
  ValidationUtils.verify(false,true,"quantity is Needed to Invoicing");
  }
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  
  var billing = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
  if(temp[3]!=""){
  billing.Click();
  billing.setText(temp[3]);
  }else{ 
  ValidationUtils.verify(false,true,"billing Price is Needed to Invoicing");
  }
  
  var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 2);
  save.Click();
  Delay(3000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  Sys.Desktop.KeyDown(0x09);
  Sys.Desktop.KeyUp(0x09);
  Delay(2000);
  var Action = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPopupPickerWidget", "");
  Action.Click();
  Sys.Process("Maconomy").Refresh();
  WorkspaceUtils.DropDownList("Invoice");

  var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
  save.Click();
  Delay(3000);



   
    }
    else{ 
      if(j==(invoices.length-1)){
//        if(checkpoint){
//      selectionBilling.Click();
//      var Action = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 3).SWTObject("Button", "")
//      if(!Action.getSelection())
//      Action.Click();
//      var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
//      save.Click();
//      Delay(3000);
//      }
//      table.Keys("[Down]");
      }
    }
}

        if(checkpoint){
      selectionBilling.Click();
      var Action = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2).SWTObject("McPlainCheckboxView", "", 3).SWTObject("Button", "")
      if(!Action.getSelection()){
      for(var u=0;u<AArray.length;u++){
      var temp1 = AArray[u].split("*");
      for(var v=0;v<arrys.length;v++){
      var temp3 = arrys[v].split("*");
      if((table.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(temp1[0])!=-1)&&(table.getItem(i).getText_2(1).OleValue.toString().trim().indexOf(temp3[1])!=-1)) 
      if((aqConvert.StrToInt(table.getItem(i).getText_2(4).OleValue.toString().trim())+aqConvert.StrToInt(table.getItem(i).getText_2(6).OleValue.toString().trim())+aqConvert.StrToInt(temp1[1].toString().trim()))==temp3[2]){
      Log.Message((aqConvert.StrToInt(table.getItem(i).getText_2(4).OleValue.toString().trim())+aqConvert.StrToInt(table.getItem(i).getText_2(6).OleValue.toString().trim())+aqConvert.StrToInt(temp1[1].toString().trim())))
      Log.Message(temp3[2])
      Log.Message((aqConvert.StrToInt(table.getItem(i).getText_2(4).OleValue.toString().trim())+aqConvert.StrToInt(table.getItem(i).getText_2(4).OleValue.toString().trim())+aqConvert.StrToInt(temp1[1].toString().trim()))==temp3[2])
      Action.Click();
      var save = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 3);
      save.Click();
      Delay(3000);
      }
      }
      }
}
  }
        if(i<(table.getItemCount()-1))
        table.Keys("[Down]");

    }
   if(table.getItemCount()>0){ 
      var approve = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 7);
      Sys.HighlightObject(approve);
      approve.Click();
      Delay(5000);
      var draftInvoice = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 9).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 8);
      Sys.HighlightObject(draftInvoice);
      draftInvoice.Click();
      Delay(5000);
      var Ref = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1);
      Ref.Refresh();
      var invoiceediting = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
      invoiceediting.Click();
      Delay(2000);
//      Ref.Refresh();
      if(Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).isVisible())
      var submitdraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 7);
      else
      var submitdraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 6)
      
      Sys.HighlightObject(submitdraft);
      submitdraft.Click();
      Delay(5000);
      var printdraft = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 2).SWTObject("SingleToolItemControl", "", 10);
      Sys.HighlightObject(printdraft);
      printdraft.Click();
      Delay(5000);
      var approvalbar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("PTabItemPanel", "", 3).SWTObject("TabControl", "");
      approvalbar.Click();
      Delay(5000);
      var Allapproval = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 7);
      Allapproval.Click();
      Delay(3000);
//      var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
     var approver_table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 6).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McTableWidget", "").SWTObject("McGrid", "", 2);
     var y=0;
     for(var z=0;z<approver_table.getItemCount();z++){ 
     approvers="";
     if(approver_table.getItem(z).getText_2(8)!="Approved"){
     
     if(approver_table.getItem(z).getText_2(4).OleValue.toString().trim()!="")
     approvers = approver_table.getItem(z).getText_2(4).OleValue.toString().trim()+"*"+approver_table.getItem(z).getText_2(3).OleValue.toString().trim();
     else
     approvers = "*"+approver_table.getItem(z).getText_2(3).OleValue.toString().trim();
     
     Log.Message(approvers)
     Log.Message(JobDetail[0])
     Log.Message(JobDetail[1])
     ReportUtils.logStep("INFO","Invoice Approver level : " +z+ " Approver :" +approvers);
     Approve_Level[y] = jobNo+"*"+JobDetail[1]+"*"+approvers;
     Log.Message(Approve_Level[y])
     y++;
     }
     }
     
    }
  
  
HRData = WorkspaceUtils.goToHR();
LoginEmp = WorkspaceUtils.Credentiallogin(Project.Path+excelName, "userRoles");

if(JobDetail[0]!="")
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,JobDetail[0]);
else
UserPasswd = WorkspaceUtils.Login_Match(Approve_Level,LoginEmp,HRData,"");

RestMaconomy(UserPasswd);
}